using System;
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

        /// <summary>
        /// Parses the container without decoding pixels; pages are decoded one at a
        /// time through <see cref="ImagePages.Decode"/>. Caller disposes the result.
        /// </summary>
        public static ImagePages Open(string path)
        {
            var bytes = File.ReadAllBytes(path);
            try
            {
                return OpenWithGdi(bytes);
            }
            catch (Exception gdiError) when (gdiError is ArgumentException || gdiError is ExternalException)
            {
                try
                {
                    return OpenWicBytes(bytes);
                }
                catch (Exception wicError)
                {
                    throw new UnreadableImageException(
                        "This file couldn't be decoded (" + gdiError.GetType().Name + ").", wicError);
                }
            }
        }

        internal static ImagePages OpenWithWic(string path)
        {
            return OpenWicBytes(File.ReadAllBytes(path));
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

        private static ImagePages OpenWithGdi(byte[] bytes)
        {
            var stream = new MemoryStream(bytes);
            Image image = null;
            try
            {
                image = Image.FromStream(stream, false, true);
                var orientation = ReadOrientation(image);
                var frames = image.RawFormat.Guid == ImageFormat.Tiff.Guid
                    ? Math.Max(1, image.GetFrameCount(FrameDimension.Page))
                    : 1;
                var source = image;
                return new ImagePages(
                    frames,
                    index =>
                    {
                        if (frames > 1) source.SelectActiveFrame(FrameDimension.Page, index);
                        var page = OnWhite(source);
                        try
                        {
                            page.RotateFlip(Upright(orientation));
                            return page;
                        }
                        catch
                        {
                            page.Dispose();
                            throw;
                        }
                    },
                    new IDisposable[] { image, stream });
            }
            catch
            {
                if (image != null) image.Dispose();
                stream.Dispose();
                throw;
            }
        }

        private static ImagePages OpenWicBytes(byte[] bytes)
        {
            var stream = new MemoryStream(bytes);
            try
            {
                var decoder = System.Windows.Media.Imaging.BitmapDecoder.Create(
                    stream,
                    System.Windows.Media.Imaging.BitmapCreateOptions.IgnoreColorProfile,
                    System.Windows.Media.Imaging.BitmapCacheOption.OnDemand);
                var frames = decoder is System.Windows.Media.Imaging.TiffBitmapDecoder
                    ? Math.Max(1, decoder.Frames.Count)
                    : 1;
                return new ImagePages(frames, index => DecodeWicFrame(decoder.Frames[index]), new IDisposable[] { stream });
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        private static Bitmap DecodeWicFrame(System.Windows.Media.Imaging.BitmapFrame frame)
        {
            var bgra = new System.Windows.Media.Imaging.FormatConvertedBitmap(
                frame, System.Windows.Media.PixelFormats.Bgra32, null, 0);
            var width = bgra.PixelWidth;
            var height = bgra.PixelHeight;
            var source = new byte[checked(width * height * 4)];
            bgra.CopyPixels(source, width * 4, 0);

            var buffer = new PixelBuffer(width, height);
            var target = buffer.Data;
            for (int s = 0, t = 0; s < source.Length; s += 4, t += 3)
            {
                var alpha = source[s + 3];
                if (alpha == 255)
                {
                    target[t] = source[s];
                    target[t + 1] = source[s + 1];
                    target[t + 2] = source[s + 2];
                }
                else
                {
                    var inverse = 255 - alpha;
                    target[t] = (byte)((source[s] * alpha + 255 * inverse + 127) / 255);
                    target[t + 1] = (byte)((source[s + 1] * alpha + 255 * inverse + 127) / 255);
                    target[t + 2] = (byte)((source[s + 2] * alpha + 255 * inverse + 127) / 255);
                }
            }

            var page = buffer.ToBitmap();
            try
            {
                page.RotateFlip(Upright(WicOrientation(frame)));
                return page;
            }
            catch
            {
                page.Dispose();
                throw;
            }
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
    /// <summary>An opened image file: a page count, and pages decoded one at a time.</summary>
    internal sealed class ImagePages : IDisposable
    {
        private readonly Func<int, Bitmap> _decode;
        private readonly IDisposable[] _owned;

        internal ImagePages(int count, Func<int, Bitmap> decode, IDisposable[] owned)
        {
            Count = count;
            _decode = decode;
            _owned = owned;
        }

        public int Count { get; }

        /// <summary>A new upright 24-bit page on white. Caller disposes it.</summary>
        public Bitmap Decode(int index)
        {
            if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
            return _decode(index);
        }

        public void Dispose()
        {
            foreach (var item in _owned) item.Dispose();
        }
    }
}
