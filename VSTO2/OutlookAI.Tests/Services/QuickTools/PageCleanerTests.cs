using System;
using System.Drawing;
using System.IO;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PageCleanerTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-cleaner", Path.GetRandomFileName());

        public PageCleanerTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private string Photo(out PointD[] cornersPx)
        {
            cornersPx = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };
            var path = Path.Combine(_dir, "upright.jpg");
            using (var page = SyntheticPages.Page(850, 1100, SyntheticPages.DimPaper, SyntheticPages.DimInk, marker: true))
            using (var photo = SyntheticPages.Photograph(page, 1600, 1200, cornersPx, SyntheticPages.Desk))
            {
                JpegFiles.Save(photo, path, JpegFiles.OriginalQuality);
            }
            return path;
        }

        private static PagePlan PlanFor(PointD[] cornersPx, int width, int height, string strength)
        {
            var corners = "[";
            for (var i = 0; i < 4; i++)
            {
                corners += (i > 0 ? "," : "") + "[" + (cornersPx[i].X / width).ToString("R", System.Globalization.CultureInfo.InvariantCulture)
                    + "," + (cornersPx[i].Y / height).ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "]";
            }
            return PagePlan.FromJson("{\"kind\":\"document_page\",\"reason\":\"\",\"corners\":" + corners
                + "],\"color\":\"grayscale\",\"strength\":\"" + strength + "\"}");
        }

        [Fact]
        public void WriteCleaned_FlattensAndCleansThePage()
        {
            PointD[] cornersPx;
            var upright = Photo(out cornersPx);
            var output = Path.Combine(_dir, "cleaned.jpg");

            PageCleaner.WriteCleaned(upright, PlanFor(cornersPx, 1600, 1200, "light"), PageColor.Grayscale, output);

            using (var cleaned = JpegFiles.Load(output))
            {
                var px = PixelBuffer.FromBitmap(cleaned);
                var marker = SyntheticPages.DarkCentroid(px, 40);
                Assert.InRange(marker.X, 0.13, 0.17);
                Assert.InRange(marker.Y, 0.118, 0.158);
                // Bottom margin of the page: paper, now near white.
                var o = ((int)(0.95 * px.Height) * px.Width + px.Width / 2) * 3;
                Assert.True(px.Data[o] > 200);
            }
        }

        [Fact]
        public void WriteCleaned_FallbackPlan_KeepsTheWholeImage()
        {
            var upright = Path.Combine(_dir, "whole.jpg");
            using (var bitmap = new Bitmap(400, 300)) JpegFiles.Save(bitmap, upright, 95);
            var output = Path.Combine(_dir, "whole-cleaned.jpg");

            PageCleaner.WriteCleaned(upright, PagePlan.Fallback(), PageColor.Grayscale, output);

            using (var cleaned = JpegFiles.Load(output))
            {
                Assert.Equal(new Size(400, 300), cleaned.Size);
            }
        }

        [Fact]
        public void WriteRotated_TurnsClockwise()
        {
            var source = Path.Combine(_dir, "wide.jpg");
            using (var bitmap = new Bitmap(40, 20)) JpegFiles.Save(bitmap, source, 95);
            var output = Path.Combine(_dir, "tall.jpg");

            PageCleaner.WriteRotated(source, 1, output);

            using (var rotated = JpegFiles.Load(output))
            {
                Assert.Equal(new Size(20, 40), rotated.Size);
            }
        }

        [Theory]
        [InlineData(0, RotateFlipType.RotateNoneFlipNone)]
        [InlineData(1, RotateFlipType.Rotate90FlipNone)]
        [InlineData(2, RotateFlipType.Rotate180FlipNone)]
        [InlineData(3, RotateFlipType.Rotate270FlipNone)]
        [InlineData(-1, RotateFlipType.Rotate270FlipNone)]
        [InlineData(5, RotateFlipType.Rotate90FlipNone)]
        public void QuarterTurns_WrapAround(int turns, RotateFlipType expected)
        {
            Assert.Equal(expected, PageCleaner.QuarterTurns(turns));
        }
    }
}
