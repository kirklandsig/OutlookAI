using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImageLoaderTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-imageloader", Path.GetRandomFileName());

        public ImageLoaderTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        // 40 x 20, left half red, right half blue.
        private static Bitmap Sample()
        {
            var bitmap = new Bitmap(40, 20, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.FillRectangle(Brushes.Red, 0, 0, 20, 20);
                g.FillRectangle(Brushes.Blue, 20, 0, 20, 20);
            }
            return bitmap;
        }

        private string SaveWithOrientation(int orientation)
        {
            var path = Path.Combine(_dir, "o" + orientation + ".jpg");
            using (var bitmap = Sample())
            {
                var item = (PropertyItem)FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                item.Id = 0x0112;
                item.Type = 3;
                item.Len = 2;
                item.Value = BitConverter.GetBytes((ushort)orientation);
                bitmap.SetPropertyItem(item);
                JpegFiles.Save(bitmap, path, 95);
            }
            return path;
        }

        [Theory]
        [InlineData(1, RotateFlipType.RotateNoneFlipNone)]
        [InlineData(2, RotateFlipType.RotateNoneFlipX)]
        [InlineData(3, RotateFlipType.Rotate180FlipNone)]
        [InlineData(4, RotateFlipType.Rotate180FlipX)]
        [InlineData(5, RotateFlipType.Rotate90FlipX)]
        [InlineData(6, RotateFlipType.Rotate90FlipNone)]
        [InlineData(7, RotateFlipType.Rotate270FlipX)]
        [InlineData(8, RotateFlipType.Rotate270FlipNone)]
        [InlineData(0, RotateFlipType.RotateNoneFlipNone)]
        public void Upright_MapsEveryExifOrientation(int orientation, RotateFlipType expected)
        {
            Assert.Equal(expected, ImageLoader.Upright(orientation));
        }

        [Fact]
        public void LoadPages_TurnsAPhoneShotUpright()
        {
            var pages = ImageLoader.LoadPages(SaveWithOrientation(6));
            try
            {
                var page = Assert.Single(pages);
                Assert.Equal(20, page.Width);
                Assert.Equal(40, page.Height);
                // Rotated 90 degrees clockwise: the red left half is now on top.
                Assert.True(page.GetPixel(10, 5).R > 200);
                Assert.True(page.GetPixel(10, 35).B > 200);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void LoadPages_MultiPageTiff_GivesEveryFrame()
        {
            var path = Path.Combine(_dir, "fax.tif");
            using (var first = Sample())
            using (var second = Sample())
            using (var third = Sample())
            {
                var tiff = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Tiff.Guid);
                using (var parameters = new EncoderParameters(1))
                {
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                    first.Save(path, tiff, parameters);
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                    first.SaveAdd(second, parameters);
                    first.SaveAdd(third, parameters);
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
                    first.SaveAdd(parameters);
                }
            }

            var pages = ImageLoader.LoadPages(path);
            try
            {
                Assert.Equal(3, pages.Count);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void TransparentPng_ComesOutOnWhite()
        {
            var path = Path.Combine(_dir, "screenshot.png");
            using (var png = new Bitmap(30, 30, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(png)) g.Clear(Color.Transparent);
                png.SetPixel(15, 15, Color.Black);
                png.Save(path, ImageFormat.Png);
            }

            var pages = ImageLoader.LoadPages(path);
            try
            {
                var pixel = pages[0].GetPixel(2, 2);
                Assert.Equal(255, pixel.R);
                Assert.Equal(255, pixel.G);
                Assert.Equal(255, pixel.B);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void LoadWithWic_DecodesWhatWindowsCan()
        {
            var path = Path.Combine(_dir, "plain.png");
            using (var bitmap = Sample()) bitmap.Save(path, ImageFormat.Png);

            var pages = ImageLoader.LoadWithWic(path);
            try
            {
                Assert.Equal(40, pages[0].Width);
                Assert.True(pages[0].GetPixel(5, 5).R > 200);
            }
            finally { foreach (var p in pages) p.Dispose(); }
        }

        [Fact]
        public void LoadPages_NotAnImage_IsUnreadable()
        {
            var path = Path.Combine(_dir, "fake.jpg");
            File.WriteAllText(path, "this is not an image");

            Assert.Throws<UnreadableImageException>(() => ImageLoader.LoadPages(path));
        }

        [Fact]
        public void JpegFiles_Scaled_CapsTheLongEdge_AndNeverEnlarges()
        {
            using (var big = new Bitmap(3200, 1600))
            using (var small = new Bitmap(300, 200))
            using (var a = JpegFiles.Scaled(big, 1600))
            using (var b = JpegFiles.Scaled(small, 1600))
            {
                Assert.Equal(new Size(1600, 800), a.Size);
                Assert.Equal(new Size(300, 200), b.Size);
            }
        }

        [Fact]
        public void JpegFiles_Load_DoesntKeepTheFileLocked()
        {
            var path = Path.Combine(_dir, "unlocked.jpg");
            using (var bitmap = Sample()) JpegFiles.Save(bitmap, path, 85);

            using (var loaded = JpegFiles.Load(path))
            {
                File.Delete(path);
                Assert.Equal(40, loaded.Width);
            }
        }

        [Fact]
        public void PixelBuffer_RoundTripsThroughABitmap()
        {
            using (var bitmap = Sample())
            {
                var buffer = PixelBuffer.FromBitmap(bitmap);
                Assert.Equal(new byte[] { 0, 0, 255 }, buffer.Data.Take(3).ToArray()); // BGR red
                using (var back = buffer.ToBitmap())
                {
                    Assert.Equal(Color.FromArgb(255, 0, 0, 255).ToArgb(), back.GetPixel(30, 10).ToArgb());
                }
            }
        }
    }
}
