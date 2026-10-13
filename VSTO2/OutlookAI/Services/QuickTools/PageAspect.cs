using System;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// Width ÷ height of the real page behind a photographed outline (TL, TR,
    /// BR, BL in pixels), assuming the optical center is the image center
    /// (Zhang &amp; He, "Whiteboard scanning and image enhancement"). Falls back to
    /// the longer opposite edges when the photo shows no usable perspective.
    /// </summary>
    internal static class PageAspect
    {
        private const double MinRatio = 0.2;
        private const double MaxRatio = 5.0;

        public static double Estimate(PointD[] c, int imageWidth, int imageHeight)
        {
            var fallback = EdgeRatio(c);
            double u0 = imageWidth / 2.0, v0 = imageHeight / 2.0;
            var m1 = V(c[0], u0, v0);
            var m2 = V(c[1], u0, v0);
            var m4 = V(c[2], u0, v0);
            var m3 = V(c[3], u0, v0);

            var d2 = Dot(Cross(m2, m4), m3);
            var d3 = Dot(Cross(m3, m4), m2);
            if (Math.Abs(d2) < 1e-9 || Math.Abs(d3) < 1e-9) return fallback;
            var c14 = Cross(m1, m4);
            var k2 = Dot(c14, m3) / d2;
            var k3 = Dot(c14, m2) / d3;
            var n2 = new[] { k2 * m2[0] - m1[0], k2 * m2[1] - m1[1], k2 * m2[2] - m1[2] };
            var n3 = new[] { k3 * m3[0] - m1[0], k3 * m3[1] - m1[1], k3 * m3[2] - m1[2] };

            var den = n2[2] * n3[2];
            if (Math.Abs(den) < 1e-12) return fallback; // no perspective: the edges are right
            var f2 = -(n2[0] * n3[0] + n2[1] * n3[1]) / den;
            var maxEdge = Math.Max(imageWidth, imageHeight);
            // A real camera's focal length is a fraction to a few times the image size.
            if (f2 <= 0 || f2 < Sq(0.25 * maxEdge) || f2 > Sq(4.0 * maxEdge)) return fallback;

            var ratio = Math.Sqrt(
                ((Sq(n2[0]) + Sq(n2[1])) / f2 + Sq(n2[2]))
                / ((Sq(n3[0]) + Sq(n3[1])) / f2 + Sq(n3[2])));
            return ratio >= MinRatio && ratio <= MaxRatio ? ratio : fallback;
        }

        public static double EdgeRatio(PointD[] c) =>
            Math.Max(Distance(c[0], c[1]), Distance(c[3], c[2]))
            / Math.Max(Distance(c[0], c[3]), Distance(c[1], c[2]));

        private static double[] V(PointD p, double u0, double v0) => new[] { p.X - u0, p.Y - v0, 1.0 };

        private static double[] Cross(double[] a, double[] b) => new[]
        {
            a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0],
        };

        private static double Dot(double[] a, double[] b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];

        private static double Sq(double v) => v * v;

        private static double Distance(PointD a, PointD b) => Math.Sqrt(Sq(a.X - b.X) + Sq(a.Y - b.Y));
    }
}
