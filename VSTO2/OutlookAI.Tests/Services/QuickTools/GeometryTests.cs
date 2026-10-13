using System;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class GeometryTests
    {
        private static PointD[] P(params double[] xy)
        {
            var points = new PointD[xy.Length / 2];
            for (var i = 0; i < points.Length; i++) points[i] = new PointD(xy[2 * i], xy[2 * i + 1]);
            return points;
        }

        [Fact]
        public void Homography_MapsTheFourPointsExactly()
        {
            var from = P(0, 0, 850, 0, 850, 1100, 0, 1100);
            var to = P(470, 140, 1150, 190, 1290, 1080, 330, 1010);

            var map = Homography.FromPoints(from, to);

            for (var i = 0; i < 4; i++)
            {
                var mapped = map.Map(from[i].X, from[i].Y);
                Assert.Equal(to[i].X, mapped.X, 6);
                Assert.Equal(to[i].Y, mapped.Y, 6);
            }
        }

        [Fact]
        public void Homography_ThereAndBack_IsTheIdentity()
        {
            var square = P(0, 0, 1, 0, 1, 1, 0, 1);
            var trapezoid = P(0.29, 0.12, 0.72, 0.16, 0.81, 0.9, 0.21, 0.84);
            var there = Homography.FromPoints(square, trapezoid);
            var back = Homography.FromPoints(trapezoid, square);

            var mid = there.Map(0.3, 0.6);
            var round = back.Map(mid.X, mid.Y);

            Assert.Equal(0.3, round.X, 9);
            Assert.Equal(0.6, round.Y, 9);
        }

        [Fact]
        public void Homography_RejectsCollinearPoints()
        {
            Assert.Throws<InvalidOperationException>(
                () => Homography.FromPoints(P(0, 0, 1, 1, 2, 2, 3, 3), P(0, 0, 1, 0, 1, 1, 0, 1)));
        }

        [Fact]
        public void Quad_ReadingOrderOutline_IsConvexClockwise()
        {
            Assert.True(Quad.IsConvexClockwise(P(0.29, 0.12, 0.72, 0.16, 0.81, 0.9, 0.21, 0.84)));
        }

        // TL, BL, BR, TR would produce a mirrored page.
        [Fact]
        public void Quad_MirroredOrder_IsRejected()
        {
            Assert.False(Quad.IsConvexClockwise(P(0, 0, 0, 1, 1, 1, 1, 0)));
        }

        [Fact]
        public void Quad_SelfCrossingOrder_IsRejected()
        {
            Assert.False(Quad.IsConvexClockwise(P(0, 0, 1, 1, 1, 0, 0, 1)));
        }

        [Fact]
        public void Quad_AreaOfTheUnitSquare_IsOne()
        {
            Assert.Equal(1.0, Quad.Area(P(0, 0, 1, 0, 1, 1, 0, 1)), 9);
        }
    }
}
