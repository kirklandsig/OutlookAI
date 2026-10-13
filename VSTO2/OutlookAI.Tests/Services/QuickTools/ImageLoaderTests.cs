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
        public void Open_TurnsAPhoneShotUpright()
        {
            using (var pages = ImageLoader.Open(SaveWithOrientation(6)))
            using (var page = pages.Decode(0))
            {
                Assert.Equal(1, pages.Count);
                Assert.Equal(20, page.Width);
                Assert.Equal(40, page.Height);
                // Rotated 90 degrees clockwise: the red left half is now on top.
                Assert.True(page.GetPixel(10, 5).R > 200);
                Assert.True(page.GetPixel(10, 35).B > 200);
            }
        }

        private string SaveTiff(string name, int frames)
        {
            var path = Path.Combine(_dir, name);
            var bitmaps = Enumerable.Range(0, frames).Select(_ => Sample()).ToList();
            try
            {
                var tiff = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Tiff.Guid);
                using (var parameters = new EncoderParameters(1))
                {
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                    bitmaps[0].Save(path, tiff, parameters);
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                    for (var i = 1; i < frames; i++) bitmaps[0].SaveAdd(bitmaps[i], parameters);
                    parameters.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
                    bitmaps[0].SaveAdd(parameters);
                }
            }
            finally { foreach (var b in bitmaps) b.Dispose(); }
            return path;
        }

        private string SaveTransparentPng()
        {
            var path = Path.Combine(_dir, "screenshot.png");
            using (var png = new Bitmap(30, 30, PixelFormat.Format32bppArgb))
            {
                using (var g = Graphics.FromImage(png)) g.Clear(Color.Transparent);
                png.SetPixel(15, 15, Color.Black);
                png.Save(path, ImageFormat.Png);
            }
            return path;
        }

        [Fact]
        public void Open_MultiPageTiff_CountsPages_AndDecodesEachOnDemand()
        {
            using (var pages = ImageLoader.Open(SaveTiff("fax.tif", 3)))
            {
                Assert.Equal(3, pages.Count);
                for (var i = 0; i < 3; i++)
                {
                    using (var page = pages.Decode(i))
                    {
                        Assert.Equal(new Size(40, 20), page.Size);
                        Assert.True(page.GetPixel(5, 5).R > 200);
                    }
                }
            }
        }

        [Fact]
        public void OpenWithWic_MultiPageTiff_CountsPages()
        {
            using (var pages = ImageLoader.OpenWithWic(SaveTiff("wicfax.tif", 3)))
            {
                Assert.Equal(3, pages.Count);
                using (var page = pages.Decode(2)) Assert.Equal(40, page.Width);
            }
        }

        [Fact]
        public void OpenWithWic_Png_IsOnePage()
        {
            var path = Path.Combine(_dir, "one.png");
            using (var bitmap = Sample()) bitmap.Save(path, ImageFormat.Png);
            using (var pages = ImageLoader.OpenWithWic(path)) Assert.Equal(1, pages.Count);
        }

        [Fact]
        public void OpenWithWic_TransparentPng_ComesOutOnWhite()
        {
            using (var pages = ImageLoader.OpenWithWic(SaveTransparentPng()))
            using (var page = pages.Decode(0))
            {
                var pixel = page.GetPixel(2, 2);
                Assert.Equal(255, pixel.R);
                Assert.Equal(255, pixel.G);
                Assert.Equal(255, pixel.B);
                Assert.True(page.GetPixel(15, 15).R < 50);
            }
        }

        [Fact]
        public void TransparentPng_ComesOutOnWhite()
        {
            using (var pages = ImageLoader.Open(SaveTransparentPng()))
            using (var page = pages.Decode(0))
            {
                var pixel = page.GetPixel(2, 2);
                Assert.Equal(255, pixel.R);
                Assert.Equal(255, pixel.G);
                Assert.Equal(255, pixel.B);
            }
        }

        [Fact]
        public void OpenWithWic_DecodesWhatWindowsCan()
        {
            var path = Path.Combine(_dir, "plain.png");
            using (var bitmap = Sample()) bitmap.Save(path, ImageFormat.Png);

            using (var pages = ImageLoader.OpenWithWic(path))
            using (var page = pages.Decode(0))
            {
                Assert.Equal(40, page.Width);
                Assert.True(page.GetPixel(5, 5).R > 200);
            }
        }

        [Fact]
        public void Open_NotAnImage_IsUnreadable()
        {
            var path = Path.Combine(_dir, "fake.jpg");
            File.WriteAllText(path, "this is not an image");

            Assert.Throws<UnreadableImageException>(() => ImageLoader.Open(path));
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
        public void JpegFiles_Load_Is24Bit()
        {
            var path = Path.Combine(_dir, "page.jpg");
            using (var bitmap = Sample()) JpegFiles.Save(bitmap, path, 85);

            using (var loaded = JpegFiles.Load(path))
            {
                Assert.Equal(PixelFormat.Format24bppRgb, loaded.PixelFormat);
                Assert.True(loaded.GetPixel(5, 5).R > 200);
            }
        }

        // 5000 x 3000, left half red, right half blue.
        private string SaveLargeJpeg(string name)
        {
            var path = Path.Combine(_dir, name);
            using (var bitmap = new Bitmap(5000, 3000, PixelFormat.Format24bppRgb))
            {
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.FillRectangle(Brushes.Red, 0, 0, 2500, 3000);
                    g.FillRectangle(Brushes.Blue, 2500, 0, 2500, 3000);
                }
                JpegFiles.Save(bitmap, path, 85);
            }
            return path;
        }

        [Fact]
        public void Open_VeryLargePhoto_DecodesAtMost4000PxOnTheLongSide()
        {
            using (var pages = ImageLoader.Open(SaveLargeJpeg("big.jpg")))
            using (var page = pages.Decode(0))
            {
                Assert.Equal(new Size(4000, 2400), page.Size);
                Assert.True(page.GetPixel(100, 1200).R > 200);
                Assert.True(page.GetPixel(3900, 1200).B > 200);
            }
        }

        [Fact]
        public void OpenWithWic_VeryLargePhoto_DecodesAtMost4000PxOnTheLongSide()
        {
            using (var pages = ImageLoader.OpenWithWic(SaveLargeJpeg("wicbig.jpg")))
            using (var page = pages.Decode(0))
            {
                Assert.Equal(new Size(4000, 2400), page.Size);
                Assert.True(page.GetPixel(100, 1200).R > 200);
                Assert.True(page.GetPixel(3900, 1200).B > 200);
            }
        }

        private string Save300Square()
        {
            var path = Path.Combine(_dir, "square.jpg");
            using (var bitmap = Sample())
            using (var square = new Bitmap(bitmap, 300, 300))
            {
                JpegFiles.Save(square, path, 85);
            }
            return path;
        }

        [Fact]
        public void Open_OverThePixelLimit_IsUnreadable()
        {
            var ex = Assert.Throws<UnreadableImageException>(() => ImageLoader.Open(Save300Square(), maxSourcePixels: 10000));

            Assert.Equal("This image is too large.", ex.Message);
        }

        [Fact]
        public void OpenWithWic_OverThePixelLimit_IsUnreadable()
        {
            var ex = Assert.Throws<UnreadableImageException>(() => ImageLoader.OpenWithWic(Save300Square(), maxSourcePixels: 10000));

            Assert.Equal("This image is too large.", ex.Message);
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
