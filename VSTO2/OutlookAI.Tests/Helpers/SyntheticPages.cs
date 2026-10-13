using System.Drawing;
using System.Drawing.Imaging;
using OutlookAI.Services.QuickTools;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>Fake document photos with known geometry (fake text only).</summary>
    internal static class SyntheticPages
    {
        public static readonly Color DimPaper = Color.FromArgb(150, 142, 110);
        public static readonly Color DimInk = Color.FromArgb(100, 94, 78);
        public static readonly Color Desk = Color.FromArgb(58, 46, 38);

        /// <summary>
        /// Text lines from 30% to ~88% of the height. With marker, a black square
        /// whose top-left is at (0.1 W, 0.1 H), 0.1 W wide. With stamp, a line of
        /// bold text in that color at (0.55 W, 0.12 H).
        /// </summary>
        public static Bitmap Page(int width, int height, Color paper, Color ink, bool marker = false, Color? stamp = null)
        {
            var page = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(page))
            using (var inkBrush = new SolidBrush(ink))
            using (var font = new Font("Times New Roman", height / 40f, GraphicsUnit.Pixel))
            {
                g.Clear(paper);
                if (marker) g.FillRectangle(Brushes.Black, width * 0.1f, height * 0.1f, width * 0.1f, width * 0.1f);
                if (stamp.HasValue)
                {
                    using (var stampBrush = new SolidBrush(stamp.Value))
                    using (var bold = new Font("Arial", height / 30f, FontStyle.Bold, GraphicsUnit.Pixel))
                    {
                        g.DrawString("RECEIVED OCT 13 2026", bold, stampBrush, width * 0.55f, height * 0.12f);
                    }
                }
                for (var line = 0; line < 12; line++)
                {
                    g.DrawString("Lease line " + (line + 1) + ": tenant John Example, 555-0187",
                        font, inkBrush, width * 0.1f, height * (0.3f + line * 0.05f));
                }
            }
            return page;
        }

        /// <summary>The page as if photographed: its corners at these canvas pixels, the rest background.</summary>
        public static Bitmap Photograph(Bitmap page, int canvasWidth, int canvasHeight, PointD[] cornersPx, Color background)
        {
            var src = PixelBuffer.FromBitmap(page);
            var toPage = Homography.FromPoints(cornersPx, new[]
            {
                new PointD(0, 0), new PointD(page.Width, 0), new PointD(page.Width, page.Height), new PointD(0, page.Height),
            });
            var dst = new PixelBuffer(canvasWidth, canvasHeight);
            for (var y = 0; y < canvasHeight; y++)
            {
                for (var x = 0; x < canvasWidth; x++)
                {
                    var p = toPage.Map(x + 0.5, y + 0.5);
                    var o = (y * canvasWidth + x) * 3;
                    if (p.X >= 0 && p.X < page.Width && p.Y >= 0 && p.Y < page.Height)
                    {
                        var i = ((int)p.Y * page.Width + (int)p.X) * 3;
                        dst.Data[o] = src.Data[i];
                        dst.Data[o + 1] = src.Data[i + 1];
                        dst.Data[o + 2] = src.Data[i + 2];
                    }
                    else
                    {
                        dst.Data[o] = background.B;
                        dst.Data[o + 1] = background.G;
                        dst.Data[o + 2] = background.R;
                    }
                }
            }
            return dst.ToBitmap();
        }

        /// <summary>Centroid of pixels darker than threshold (all channels), as fractions of the image.</summary>
        public static PointD DarkCentroid(PixelBuffer px, int threshold)
        {
            double sx = 0, sy = 0;
            long n = 0;
            for (var y = 0; y < px.Height; y++)
            {
                for (var x = 0; x < px.Width; x++)
                {
                    var i = (y * px.Width + x) * 3;
                    if (px.Data[i] < threshold && px.Data[i + 1] < threshold && px.Data[i + 2] < threshold)
                    {
                        sx += x;
                        sy += y;
                        n++;
                    }
                }
            }
            return n == 0 ? new PointD(double.NaN, double.NaN) : new PointD(sx / n / px.Width, sy / n / px.Height);
        }
    }
}
