using System;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class LookProfile
    {
        public LookProfile(bool evenLight, double blackPercentile, double whitePercentile, double gamma, double sharpen)
        {
            EvenLight = evenLight;
            BlackPercentile = blackPercentile;
            WhitePercentile = whitePercentile;
            Gamma = gamma;
            Sharpen = sharpen;
        }

        public bool EvenLight { get; }
        public double BlackPercentile { get; }
        public double WhitePercentile { get; }
        /// <summary>Below 1 darkens midtones: out = x^(1/Gamma).</summary>
        public double Gamma { get; }
        public double Sharpen { get; }
    }

    /// <summary>
    /// Deterministic document cleanup: (color only) white balance on the paper,
    /// light evening, levels from the luminance's black and white points, a light
    /// unsharp mask.
    /// </summary>
    internal static class PageLook
    {
        public static readonly LookProfile Light = new LookProfile(false, 0.5, 99.5, 1.00, 0.3);
        public static readonly LookProfile Normal = new LookProfile(true, 1, 98, 0.95, 0.5);
        public static readonly LookProfile Strong = new LookProfile(true, 2, 95, 0.85, 0.8);

        private const int Block = 8;
        private const int SharpenThreshold = 2;
        // Levels only stretch when the page has at least this much tonal range; less is paper noise.
        private const int MinContrast = 48;
        // Background estimates are floored at this share of the paper so large dark areas are not whitened.
        private const float BackgroundFloor = 0.6f;

        public static LookProfile For(CleanupStrength strength)
        {
            switch (strength)
            {
                case CleanupStrength.Light: return Light;
                case CleanupStrength.Strong: return Strong;
                default: return Normal;
            }
        }

        public static void Apply(PixelBuffer px, LookProfile profile, bool grayscale)
        {
            var n = px.Width * px.Height;
            var d = px.Data;
            if (!grayscale) WhiteBalance(px);
            var lum = Luminance(px);

            byte[][] channels;
            if (grayscale)
            {
                channels = new[] { lum };
            }
            else
            {
                channels = new[] { new byte[n], new byte[n], new byte[n] };
                for (var i = 0; i < n; i++)
                {
                    channels[0][i] = d[3 * i];
                    channels[1][i] = d[3 * i + 1];
                    channels[2][i] = d[3 * i + 2];
                }
            }

            if (profile.EvenLight) EvenLight(channels, lum, px.Width, px.Height, grayscale);

            var black = Percentile(lum, profile.BlackPercentile);
            var white = Percentile(lum, profile.WhitePercentile);
            if (white - black >= MinContrast)
            {
                var lut = Levels(black, white, profile.Gamma);
                foreach (var c in channels)
                {
                    for (var i = 0; i < n; i++) c[i] = lut[c[i]];
                }
            }

            if (profile.Sharpen > 0)
            {
                foreach (var c in channels) Sharpen(c, px.Width, px.Height, profile.Sharpen);
            }

            for (var i = 0; i < n; i++)
            {
                if (grayscale)
                {
                    d[3 * i] = d[3 * i + 1] = d[3 * i + 2] = lum[i];
                }
                else
                {
                    d[3 * i] = channels[0][i];
                    d[3 * i + 1] = channels[1][i];
                    d[3 * i + 2] = channels[2][i];
                }
            }
        }

        internal static byte[] Luminance(PixelBuffer px)
        {
            var n = px.Width * px.Height;
            var d = px.Data;
            var lum = new byte[n];
            for (var i = 0; i < n; i++)
            {
                lum[i] = (byte)((114 * d[3 * i] + 587 * d[3 * i + 1] + 299 * d[3 * i + 2] + 500) / 1000);
            }
            return lum;
        }

        internal static int Percentile(byte[] values, double percent)
        {
            var histogram = new long[256];
            foreach (var v in values) histogram[v]++;
            var target = (long)Math.Ceiling(values.Length * percent / 100.0);
            if (target < 1) target = 1;
            long seen = 0;
            for (var v = 0; v < 256; v++)
            {
                seen += histogram[v];
                if (seen >= target) return v;
            }
            return 255;
        }

        // Scales each channel so the paper (median of the brightest 10%) becomes neutral.
        private static void WhiteBalance(PixelBuffer px)
        {
            var lum = Luminance(px);
            var threshold = Percentile(lum, 90);
            var hb = new long[256];
            var hg = new long[256];
            var hr = new long[256];
            long count = 0;
            var d = px.Data;
            for (var i = 0; i < lum.Length; i++)
            {
                if (lum[i] < threshold) continue;
                hb[d[3 * i]]++;
                hg[d[3 * i + 1]]++;
                hr[d[3 * i + 2]]++;
                count++;
            }
            if (count == 0) return;
            double mb = Median(hb, count), mg = Median(hg, count), mr = Median(hr, count);
            var paper = 0.114 * mb + 0.587 * mg + 0.299 * mr;
            var lutB = Gain(paper / Math.Max(mb, 1));
            var lutG = Gain(paper / Math.Max(mg, 1));
            var lutR = Gain(paper / Math.Max(mr, 1));
            for (var i = 0; i < lum.Length; i++)
            {
                d[3 * i] = lutB[d[3 * i]];
                d[3 * i + 1] = lutG[d[3 * i + 1]];
                d[3 * i + 2] = lutR[d[3 * i + 2]];
            }
        }

        private static double Median(long[] histogram, long count)
        {
            long seen = 0;
            for (var v = 0; v < 256; v++)
            {
                seen += histogram[v];
                if (seen * 2 >= count) return v;
            }
            return 255;
        }

        private static byte[] Gain(double gain)
        {
            gain = Math.Max(0.5, Math.Min(2.0, gain));
            var lut = new byte[256];
            for (var v = 0; v < 256; v++) lut[v] = Clamp((float)(v * gain));
            return lut;
        }

        // Divides by the paper brightness around each pixel: the max of each 8x8
        // block (text is darker than paper), box-blurred with a radius of 1/20 of
        // the short side, bilinear between blocks.
        private static void EvenLight(byte[][] channels, byte[] lum, int w, int h, bool lumIsChannel)
        {
            var bw = (w + Block - 1) / Block;
            var bh = (h + Block - 1) / Block;
            var background = new float[bw * bh];
            for (var by = 0; by < bh; by++)
            {
                for (var bx = 0; bx < bw; bx++)
                {
                    var max = 0;
                    var yEnd = Math.Min(h, (by + 1) * Block);
                    var xEnd = Math.Min(w, (bx + 1) * Block);
                    for (var y = by * Block; y < yEnd; y++)
                    {
                        for (var x = bx * Block; x < xEnd; x++)
                        {
                            if (lum[y * w + x] > max) max = lum[y * w + x];
                        }
                    }
                    background[by * bw + bx] = max;
                }
            }
            BoxBlur(background, bw, bh, Math.Max(1, (int)Math.Round(Math.Min(w, h) / 20.0 / Block)));

            var paper = Percentile(lum, 90);
            var floor = Math.Max(16f, BackgroundFloor * paper);
            for (var y = 0; y < h; y++)
            {
                var gy = (y + 0.5) / Block - 0.5;
                for (var x = 0; x < w; x++)
                {
                    var gx = (x + 0.5) / Block - 0.5;
                    var gain = 255f / Math.Max(Bilinear(background, bw, bh, gx, gy), floor);
                    var i = y * w + x;
                    foreach (var c in channels) c[i] = Clamp(c[i] * gain);
                    if (!lumIsChannel) lum[i] = Clamp(lum[i] * gain);
                }
            }
        }

        private static void BoxBlur(float[] a, int w, int h, int r)
        {
            var tmp = new float[a.Length];
            for (var y = 0; y < h; y++)
            {
                for (var x = 0; x < w; x++)
                {
                    float sum = 0;
                    var count = 0;
                    for (var k = -r; k <= r; k++)
                    {
                        var xx = x + k;
                        if (xx < 0 || xx >= w) continue;
                        sum += a[y * w + xx];
                        count++;
                    }
                    tmp[y * w + x] = sum / count;
                }
            }
            for (var y = 0; y < h; y++)
            {
                for (var x = 0; x < w; x++)
                {
                    float sum = 0;
                    var count = 0;
                    for (var k = -r; k <= r; k++)
                    {
                        var yy = y + k;
                        if (yy < 0 || yy >= h) continue;
                        sum += tmp[yy * w + x];
                        count++;
                    }
                    a[y * w + x] = sum / count;
                }
            }
        }

        private static float Bilinear(float[] g, int w, int h, double x, double y)
        {
            if (x < 0) x = 0; else if (x > w - 1) x = w - 1;
            if (y < 0) y = 0; else if (y > h - 1) y = h - 1;
            var x0 = (int)x;
            var y0 = (int)y;
            var x1 = Math.Min(x0 + 1, w - 1);
            var y1 = Math.Min(y0 + 1, h - 1);
            var fx = (float)(x - x0);
            var fy = (float)(y - y0);
            var top = g[y0 * w + x0] + (g[y0 * w + x1] - g[y0 * w + x0]) * fx;
            var bottom = g[y1 * w + x0] + (g[y1 * w + x1] - g[y1 * w + x0]) * fx;
            return top + (bottom - top) * fy;
        }

        private static byte[] Levels(int black, int white, double gamma)
        {
            var lut = new byte[256];
            for (var v = 0; v < 256; v++)
            {
                var x = (v - black) / (double)(white - black);
                if (x < 0) x = 0; else if (x > 1) x = 1;
                lut[v] = (byte)Math.Round(255 * Math.Pow(x, 1 / gamma));
            }
            return lut;
        }

        // Unsharp mask: v + amount * (v - 3x3 box blur), ignoring differences of 2 or less.
        private static void Sharpen(byte[] c, int w, int h, double amount)
        {
            var blur = new byte[c.Length];
            for (var y = 0; y < h; y++)
            {
                for (var x = 0; x < w; x++)
                {
                    var sum = 0;
                    for (var dy = -1; dy <= 1; dy++)
                    {
                        var yy = Math.Min(h - 1, Math.Max(0, y + dy));
                        for (var dx = -1; dx <= 1; dx++)
                        {
                            var xx = Math.Min(w - 1, Math.Max(0, x + dx));
                            sum += c[yy * w + xx];
                        }
                    }
                    blur[y * w + x] = (byte)((sum + 4) / 9);
                }
            }
            for (var i = 0; i < c.Length; i++)
            {
                var diff = c[i] - blur[i];
                if (Math.Abs(diff) > SharpenThreshold) c[i] = Clamp((float)(c[i] + amount * diff));
            }
        }

        private static byte Clamp(float v) => v <= 0 ? (byte)0 : v >= 255 ? (byte)255 : (byte)(v + 0.5f);
    }
}
