using System;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// The projective map that sends four points onto four others, e.g. an
    /// output rectangle onto a photographed page's corners.
    /// </summary>
    internal sealed class Homography
    {
        private readonly double[] _h;

        private Homography(double[] h)
        {
            _h = h;
        }

        /// <summary>
        /// a..h and 1, row-major: x' = (a x + b y + c) / (g x + h y + 1),
        /// y' = (d x + e y + f) / (g x + h y + 1).
        /// </summary>
        public double[] Coefficients => (double[])_h.Clone();

        public static Homography FromPoints(PointD[] from, PointD[] to)
        {
            if (from == null || to == null || from.Length != 4 || to.Length != 4)
                throw new ArgumentException("Four point pairs are needed.");

            var a = new double[8, 9];
            for (var i = 0; i < 4; i++)
            {
                double x = from[i].X, y = from[i].Y, u = to[i].X, v = to[i].Y;
                var r = 2 * i;
                a[r, 0] = x; a[r, 1] = y; a[r, 2] = 1; a[r, 6] = -u * x; a[r, 7] = -u * y; a[r, 8] = u;
                a[r + 1, 3] = x; a[r + 1, 4] = y; a[r + 1, 5] = 1; a[r + 1, 6] = -v * x; a[r + 1, 7] = -v * y; a[r + 1, 8] = v;
            }

            // Gauss-Jordan elimination with partial pivoting; column 8 is the right-hand side.
            for (var col = 0; col < 8; col++)
            {
                var pivot = col;
                for (var row = col + 1; row < 8; row++)
                {
                    if (Math.Abs(a[row, col]) > Math.Abs(a[pivot, col])) pivot = row;
                }
                if (Math.Abs(a[pivot, col]) < 1e-12)
                    throw new InvalidOperationException("The points don't define a projective map.");
                if (pivot != col)
                {
                    for (var k = 0; k < 9; k++)
                    {
                        var t = a[col, k];
                        a[col, k] = a[pivot, k];
                        a[pivot, k] = t;
                    }
                }
                var p = a[col, col];
                for (var k = col; k < 9; k++) a[col, k] /= p;
                for (var row = 0; row < 8; row++)
                {
                    if (row == col) continue;
                    var f = a[row, col];
                    if (f == 0) continue;
                    for (var k = col; k < 9; k++) a[row, k] -= f * a[col, k];
                }
            }

            var h = new double[9];
            for (var i = 0; i < 8; i++) h[i] = a[i, 8];
            h[8] = 1;
            return new Homography(h);
        }

        public PointD Map(double x, double y)
        {
            var w = _h[6] * x + _h[7] * y + _h[8];
            return new PointD((_h[0] * x + _h[1] * y + _h[2]) / w, (_h[3] * x + _h[4] * y + _h[5]) / w);
        }
    }
}
