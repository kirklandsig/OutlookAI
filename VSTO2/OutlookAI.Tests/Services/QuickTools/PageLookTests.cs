using System.Drawing;
using System.Linq;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PageLookTests
    {
        private static PixelBuffer DimPage(Color? stamp = null)
        {
            using (var page = SyntheticPages.Page(600, 800, SyntheticPages.DimPaper, SyntheticPages.DimInk, stamp: stamp))
            {
                return PixelBuffer.FromBitmap(page);
            }
        }

        private static byte[] Region(PixelBuffer px, double x0, double y0, double x1, double y1)
        {
            var lum = PageLook.Luminance(px);
            return Enumerable.Range((int)(y0 * px.Height), (int)((y1 - y0) * px.Height))
                .SelectMany(y => Enumerable.Range((int)(x0 * px.Width), (int)((x1 - x0) * px.Width)).Select(x => lum[y * px.Width + x]))
                .ToArray();
        }

        [Fact]
        public void Profiles_AreTheSpecDefaults()
        {
            Assert.False(PageLook.For(CleanupStrength.Light).EvenLight);
            Assert.Equal(98.0, PageLook.For(CleanupStrength.Normal).WhitePercentile);
            Assert.Equal(0.85, PageLook.For(CleanupStrength.Strong).Gamma);
            Assert.Equal(0.8, PageLook.For(CleanupStrength.Strong).Sharpen);
        }

        [Fact]
        public void DimTintedPage_GetsAWhiteBackground_AndDarkText()
        {
            var px = DimPage();

            PageLook.Apply(px, PageLook.Normal, grayscale: true);

            Assert.True(PageLook.Percentile(Region(px, 0.1, 0.9, 0.9, 0.97), 50) > 235, "background not white");
            Assert.True(PageLook.Percentile(Region(px, 0.1, 0.3, 0.9, 0.88), 1) < 60, "text not dark");
        }

        [Fact]
        public void Grayscale_IsGray()
        {
            var px = DimPage(Color.FromArgb(200, 40, 40));

            PageLook.Apply(px, PageLook.Normal, grayscale: true);

            for (var i = 0; i < px.Data.Length; i += 3)
            {
                Assert.True(px.Data[i] == px.Data[i + 1] && px.Data[i + 1] == px.Data[i + 2]);
            }
        }

        [Fact]
        public void ColorMode_KeepsARedStampRed_AndTheTintNeutral()
        {
            var px = DimPage(Color.FromArgb(200, 40, 40));

            PageLook.Apply(px, PageLook.Normal, grayscale: false);

            var red = 0;
            for (var i = 0; i < px.Data.Length; i += 3)
            {
                if (px.Data[i + 2] - px.Data[i + 1] > 60) red++;
            }
            Assert.True(red > 50, "stamp lost its color");
            // Paper: neutral white, not yellow.
            var o = ((int)(0.95 * px.Height) * px.Width + px.Width / 2) * 3;
            Assert.True(px.Data[o] > 230 && px.Data[o + 1] > 230 && px.Data[o + 2] > 230);
        }

        [Fact]
        public void FlatImage_StaysUniform()
        {
            var px = new PixelBuffer(200, 200);
            for (var i = 0; i < px.Data.Length; i++) px.Data[i] = 120;

            PageLook.Apply(px, PageLook.Strong, grayscale: true);

            Assert.True(px.Data.All(v => v == px.Data[0]));
        }

        [Fact]
        public void LargeDarkArea_StaysDark()
        {
            using (var page = SyntheticPages.Page(600, 800, SyntheticPages.DimPaper, SyntheticPages.DimInk))
            {
                using (var g = Graphics.FromImage(page))
                using (var fill = new SolidBrush(Color.FromArgb(40, 40, 40)))
                {
                    g.FillRectangle(fill, 60, 280, 210, 200);
                }
                var px = PixelBuffer.FromBitmap(page);

                PageLook.Apply(px, PageLook.Normal, grayscale: true);

                Assert.True(PageLook.Percentile(Region(px, 0.135, 0.375, 0.415, 0.575), 50) < 100);
            }
        }

        [Fact]
        public void SparsePage_PaperStaysClean()
        {
            var px = new PixelBuffer(600, 800);
            for (var y = 0; y < 800; y++)
            {
                for (var x = 0; x < 600; x++)
                {
                    var v = (byte)(243 + ((x * 7 + y * 13) % 25) - 12);
                    var o = (y * 600 + x) * 3;
                    px.Data[o] = px.Data[o + 1] = px.Data[o + 2] = v;
                }
            }
            foreach (var y in new[] { 200, 400, 600 })
            {
                for (var x = 100; x < 300; x++)
                {
                    var o = (y * 600 + x) * 3;
                    px.Data[o] = px.Data[o + 1] = px.Data[o + 2] = 60;
                }
            }

            PageLook.Apply(px, PageLook.Normal, grayscale: true);

            Assert.True(PageLook.Percentile(Region(px, 0.05, 0.05, 0.95, 0.2), 1) > 225);
        }

        [Fact]
        public void Apply_IsDeterministic()
        {
            var a = DimPage();
            var b = a.Clone();

            PageLook.Apply(a, PageLook.Strong, grayscale: false);
            PageLook.Apply(b, PageLook.Strong, grayscale: false);

            Assert.Equal(a.Data, b.Data);
        }
    }
}
