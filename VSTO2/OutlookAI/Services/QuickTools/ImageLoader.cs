using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class UnreadableImageException : Exception
    {
        public UnreadableImageException(string message, Exception inner) : base(message, inner)
        {
        }
    }

    /// <summary>
    /// Decodes an image file into upright 24-bit pages on white: EXIF orientation
    /// applied, one page per TIFF frame, a GIF's first frame. GDI+ first; WIC
    /// (WebP, HEIC, ... when Windows has the codec) when GDI+ can't.
    /// </summary>
    internal static class ImageLoader
    {
        private const int ExifOrientation = 0x0112;

        /// <summary>Caller disposes the bitmaps.</summary>
        public static List<Bitmap> LoadPages(string path)
        {
            try
            {
                return LoadWithGdi(path);
            }
            catch (Exception gdiError) when (gdiError is ArgumentException || gdiError is ExternalException)
            {
                try
                {
                    return LoadWithWic(path);
                }
                catch (Exception wicError)
                {
                    throw new UnreadableImageException("This file couldn't be decoded.", wicError);
                }
            }
        }

        internal static RotateFlipType Upright(int orientation)
        {
            switch (orientation)
            {
                case 2: return RotateFlipType.RotateNoneFlipX;
                case 3: return RotateFlipType.Rotate180FlipNone;
                case 4: return RotateFlipType.Rotate180FlipX;
                case 5: return RotateFlipType.Rotate90FlipX;
                case 6: return RotateFlipType.Rotate90FlipNone;
                case 7: return RotateFlipType.Rotate270FlipX;
                case 8: return RotateFlipType.Rotate270FlipNone;
                default: return RotateFlipType.RotateNoneFlipNone;
            }
        }

        private static List<Bitmap> LoadWithGdi(string path)
        {
            var pages = new List<Bitmap>();
            using (var stream = new MemoryStream(File.ReadAllBytes(path)))
            using (var image = Image.FromStream(stream, false, true))
            {
                var orientation = ReadOrientation(image);
                var frames = image.RawFormat.Guid == ImageFormat.Tiff.Guid
                    ? image.GetFrameCount(FrameDimension.Page)
                    : 1;
                try
                {
                    for (var i = 0; i < frames; i++)
                    {
                        if (frames > 1) image.SelectActiveFrame(FrameDimension.Page, i);
                        var page = OnWhite(image);
                        page.RotateFlip(Upright(orientation));
                        pages.Add(page);
                    }
                }
                catch
                {
                    foreach (var p in pages) p.Dispose();
                    throw;
                }
            }
            return pages;
        }

        internal static List<Bitmap> LoadWithWic(string path)
        {
            var pages = new List<Bitmap>();
            using (var stream = new MemoryStream(File.ReadAllBytes(path)))
            {
                var decoder = System.Windows.Media.Imaging.BitmapDecoder.Create(
                    stream,
                    System.Windows.Media.Imaging.BitmapCreateOptions.IgnoreColorProfile,
                    System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
                try
                {
                    foreach (var frame in decoder.Frames)
                    {
                        var bgr = new System.Windows.Media.Imaging.FormatConvertedBitmap(
                            frame, System.Windows.Media.PixelFormats.Bgr24, null, 0);
                        var buffer = new PixelBuffer(bgr.PixelWidth, bgr.PixelHeight);
                        bgr.CopyPixels(buffer.Data, buffer.Width * 3, 0);
                        var page = buffer.ToBitmap();
                        page.RotateFlip(Upright(WicOrientation(frame)));
                        pages.Add(page);
                    }
                }
                catch
                {
                    foreach (var p in pages) p.Dispose();
                    throw;
                }
            }
            return pages;
        }

        private static int ReadOrientation(Image image)
        {
            if (Array.IndexOf(image.PropertyIdList, ExifOrientation) < 0) return 1;
            var value = image.GetPropertyItem(ExifOrientation).Value;
            return value != null && value.Length >= 2 ? BitConverter.ToUInt16(value, 0) : 1;
        }

        private static int WicOrientation(System.Windows.Media.Imaging.BitmapFrame frame)
        {
            try
            {
                var metadata = frame.Metadata as System.Windows.Media.Imaging.BitmapMetadata;
                var value = metadata == null ? null : metadata.GetQuery("System.Photo.Orientation");
                return value is ushort ? (ushort)value : 1;
            }
            catch (NotSupportedException)
            {
                return 1;
            }
            catch (InvalidOperationException)
            {
                return 1;
            }
        }

        // Normalizes any format (palette, alpha, CMYK, 16-bit) to 24-bit on white.
        private static Bitmap OnWhite(Image image)
        {
            var page = new Bitmap(image.Width, image.Height, PixelFormat.Format24bppRgb);
            page.SetResolution(96, 96);
            using (var g = Graphics.FromImage(page))
            {
                g.Clear(Color.White);
                g.InterpolationMode = InterpolationMode.NearestNeighbor;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height));
            }
            return page;
        }
    }
}
