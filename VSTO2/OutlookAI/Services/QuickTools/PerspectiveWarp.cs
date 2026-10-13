using System;
using System.Drawing;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>Maps a photographed page (4 corners in pixels) onto an upright rectangle.</summary>
    internal static class PerspectiveWarp
    {
        public const int MaxLongEdge = 3300;
        public const int MaxShortEdge = 2550;

        /// <summary>
        /// Height from the longer of the left and right edges, width from the
        /// page's estimated aspect ratio; scaled down (never up) to fit 2550 x 3300
        /// (or 3300 x 2550 when wider than tall).
        /// </summary>
        public static Size OutputSize(PointD[] cornersPx, int imageWidth, int imageHeight)
        {
            var height = Math.Max(Distance(cornersPx[0], cornersPx[3]), Distance(cornersPx[1], cornersPx[2]));
            var width = height * PageAspect.Estimate(cornersPx, imageWidth, imageHeight);
            double maxW = width >= height ? MaxLongEdge : MaxShortEdge;
            double maxH = width >= height ? MaxShortEdge : MaxLongEdge;
            var scale = Math.Min(1.0, Math.Min(maxW / width, maxH / height));
            return new Size(Math.Max(1, (int)Math.Round(width * scale)), Math.Max(1, (int)Math.Round(height * scale)));
        }

        public static PixelBuffer Warp(PixelBuffer src, PointD[] cornersPx)
        {
            var size = OutputSize(cornersPx, src.Width, src.Height);
            int outW = size.Width, outH = size.Height;
            var map = Homography.FromPoints(
                new[] { new PointD(0, 0), new PointD(outW, 0), new PointD(outW, outH), new PointD(0, outH) },
                cornersPx);
            var c = map.Coefficients;
            var dst = new PixelBuffer(outW, outH);
            var s = src.Data;
            var d = dst.Data;
            for (var y = 0; y < outH; y++)
            {
                var fy = y + 0.5;
                // Along a row the numerators and the denominator are linear in x.
                var nx = c[0] * 0.5 + c[1] * fy + c[2];
                var ny = c[3] * 0.5 + c[4] * fy + c[5];
                var nw = c[6] * 0.5 + c[7] * fy + c[8];
                var o = y * outW * 3;
                for (var x = 0; x < outW; x++, o += 3)
                {
                    Sample(s, src.Width, src.Height, nx / nw - 0.5, ny / nw - 0.5, d, o);
                    nx += c[0];
                    ny += c[3];
                    nw += c[6];
                }
            }
            return dst;
        }

        private static void Sample(byte[] s, int w, int h, double x, double y, byte[] d, int o)
        {
            if (x < 0) x = 0; else if (x > w - 1) x = w - 1;
            if (y < 0) y = 0; else if (y > h - 1) y = h - 1;
            var x0 = (int)x;
            var y0 = (int)y;
            var x1 = x0 + 1 < w ? x0 + 1 : x0;
            var y1 = y0 + 1 < h ? y0 + 1 : y0;
            var fx = x - x0;
            var fy = y - y0;
            var i00 = (y0 * w + x0) * 3;
            var i10 = (y0 * w + x1) * 3;
            var i01 = (y1 * w + x0) * 3;
            var i11 = (y1 * w + x1) * 3;
            for (var k = 0; k < 3; k++)
            {
                var top = s[i00 + k] + (s[i10 + k] - s[i00 + k]) * fx;
                var bottom = s[i01 + k] + (s[i11 + k] - s[i01 + k]) * fx;
                d[o + k] = (byte)(top + (bottom - top) * fy + 0.5);
            }
        }

        private static double Distance(PointD a, PointD b) =>
            Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));
    }
}
