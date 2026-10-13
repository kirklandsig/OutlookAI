using System;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PerspectiveWarpTests
    {
        // A 0.85 x 1.10 page (Letter, portrait) seen by a pinhole camera.
        private static PointD[] Projected(double ax, double ay, double az, double distance, int width, int height, double focal)
        {
            double cx = Math.Cos(ax), sx = Math.Sin(ax), cy = Math.Cos(ay), sy = Math.Sin(ay), cz = Math.Cos(az), sz = Math.Sin(az);
            double[,] r =
            {
                { cz * cy, cz * sy * sx - sz * cx, cz * sy * cx + sz * sx },
                { sz * cy, sz * sy * sx + cz * cx, sz * sy * cx - cz * sx },
                { -sy, cy * sx, cy * cx },
            };
            var corners = new[] { new[] { -0.425, -0.55 }, new[] { 0.425, -0.55 }, new[] { 0.425, 0.55 }, new[] { -0.425, 0.55 } };
            var result = new PointD[4];
            for (var i = 0; i < 4; i++)
            {
                double px = corners[i][0], py = corners[i][1];
                var x = r[0, 0] * px + r[0, 1] * py + 0.05;
                var y = r[1, 0] * px + r[1, 1] * py - 0.03;
                var z = r[2, 0] * px + r[2, 1] * py + distance;
                result[i] = new PointD(focal * x / z + width / 2.0, focal * y / z + height / 2.0);
            }
            return result;
        }

        [Theory]
        [InlineData(0.35, 0.2, 0.05, 3.0)]
        [InlineData(-0.5, 0.1, -0.1, 2.5)]
        [InlineData(0.6, -0.4, 0.2, 3.5)]
        public void PageAspect_RecoversTheTrueProportions_OfAnAngledPhoto(double ax, double ay, double az, double distance)
        {
            var corners = Projected(ax, ay, az, distance, 4032, 3024, 3000);

            Assert.Equal(0.85 / 1.10, PageAspect.Estimate(corners, 4032, 3024), 2);
            // The longer-edges rule is visibly off for these.
            Assert.True(Math.Abs(PageAspect.EdgeRatio(corners) - 0.85 / 1.10) > 0.04);
        }

        [Fact]
        public void PageAspect_NoPerspective_UsesTheEdges()
        {
            var corners = new[] { new PointD(100, 100), new PointD(950, 100), new PointD(950, 1200), new PointD(100, 1200) };

            Assert.Equal(850.0 / 1100.0, PageAspect.Estimate(corners, 1600, 1600), 6);
        }

        [Fact]
        public void PageAspect_ImplausibleCamera_UsesTheEdges()
        {
            // An outline no real camera produces (focal length ~10x the image).
            var corners = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };

            Assert.Equal(PageAspect.EdgeRatio(corners), PageAspect.Estimate(corners, 1600, 1200), 9);
        }

        [Fact]
        public void OutputSize_CapsHugePhotos()
        {
            var full = new[] { new PointD(0, 0), new PointD(6000, 0), new PointD(6000, 4500), new PointD(0, 4500) };

            var size = PerspectiveWarp.OutputSize(full, 6000, 4500);

            Assert.True(size.Width <= PerspectiveWarp.MaxLongEdge && size.Height <= PerspectiveWarp.MaxShortEdge);
            Assert.Equal(3300, size.Width);
        }

        [Fact]
        public void Warp_FlattensAPhotographedPage_KeepingItsLayout()
        {
            var corners = new[] { new PointD(470, 140), new PointD(1150, 190), new PointD(1290, 1080), new PointD(330, 1010) };
            using (var page = SyntheticPages.Page(850, 1100, SyntheticPages.DimPaper, SyntheticPages.DimInk, marker: true))
            using (var photo = SyntheticPages.Photograph(page, 1600, 1200, corners, SyntheticPages.Desk))
            {
                var flat = PerspectiveWarp.Warp(PixelBuffer.FromBitmap(photo), corners);

                // The black marker sits where it does on the page: centre (0.15, 0.1 + 0.05 * 850/1100).
                var marker = SyntheticPages.DarkCentroid(flat, 40);
                Assert.InRange(marker.X, 0.13, 0.17);
                Assert.InRange(marker.Y, 0.118, 0.158);
                // No desk left at the edges: the corners are paper-colored.
                var corner = (2 * flat.Width + 2) * 3;
                Assert.InRange((int)flat.Data[corner + 2], SyntheticPages.DimPaper.R - 20, SyntheticPages.DimPaper.R + 20);
            }
        }
    }
}
