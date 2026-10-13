using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace OutlookAI.Services.QuickTools
{
    internal static class JpegFiles
    {
        public const long PageQuality = 85;
        public const long OriginalQuality = 95;

        private static readonly ImageCodecInfo Codec =
            ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);

        public static void Save(Image image, string path, long quality)
        {
            using (var parameters = Quality(quality))
            {
                image.Save(path, Codec, parameters);
            }
        }

        public static byte[] Encode(Image image, long quality)
        {
            using (var parameters = Quality(quality))
            using (var stream = new MemoryStream())
            {
                image.Save(stream, Codec, parameters);
                return stream.ToArray();
            }
        }

        /// <summary>Reads the file fully first, so it isn't kept locked.</summary>
        public static Bitmap Load(string path)
        {
            using (var stream = new MemoryStream(File.ReadAllBytes(path)))
            using (var image = Image.FromStream(stream))
            {
                return new Bitmap(image);
            }
        }

        /// <summary>A copy at most maxEdge pixels on its long side (never enlarged).</summary>
        public static Bitmap Scaled(Image image, int maxEdge)
        {
            var scale = Math.Min(1.0, (double)maxEdge / Math.Max(image.Width, image.Height));
            var width = Math.Max(1, (int)Math.Round(image.Width * scale));
            var height = Math.Max(1, (int)Math.Round(image.Height * scale));
            var copy = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(copy))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(image, new Rectangle(0, 0, width, height));
            }
            return copy;
        }

        public static Bitmap LoadScaled(string path, int maxEdge)
        {
            using (var full = Load(path))
            {
                return Scaled(full, maxEdge);
            }
        }

        private static EncoderParameters Quality(long quality)
        {
            var parameters = new EncoderParameters(1);
            parameters.Param[0] = new EncoderParameter(Encoder.Quality, quality);
            return parameters;
        }
    }
}
