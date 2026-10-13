using System;

namespace OutlookAI.Services.QuickTools
{
    internal struct PointD
    {
        public PointD(double x, double y)
        {
            X = x;
            Y = y;
        }

        public double X { get; }
        public double Y { get; }
    }

    /// <summary>Checks on a page outline given as TL, TR, BR, BL.</summary>
    internal static class Quad
    {
        public static double Area(PointD[] q)
        {
            double sum = 0;
            for (var i = 0; i < 4; i++)
            {
                var a = q[i];
                var b = q[(i + 1) % 4];
                sum += a.X * b.Y - b.X * a.Y;
            }
            return Math.Abs(sum) / 2;
        }

        /// <summary>
        /// True when the outline turns the same way at every corner, clockwise on
        /// screen (y grows downward), as TL, TR, BR, BL in reading order does. A
        /// mirrored or self-crossing order fails, so it can't flip the page.
        /// </summary>
        public static bool IsConvexClockwise(PointD[] q)
        {
            for (var i = 0; i < 4; i++)
            {
                var a = q[i];
                var b = q[(i + 1) % 4];
                var c = q[(i + 2) % 4];
                var cross = (b.X - a.X) * (c.Y - b.Y) - (b.Y - a.Y) * (c.X - b.X);
                if (cross <= 0) return false;
            }
            return true;
        }
    }
}
