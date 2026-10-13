using System.Drawing;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>
    /// File in, JPEG file out, so nothing full-size outlives the call (32-bit
    /// Outlook has 2 GB of address space).
    /// </summary>
    internal static class PageCleaner
    {
        public static void WriteCleaned(string uprightPath, PagePlan plan, PageColor color, string outputPath)
        {
            PixelBuffer src;
            using (var bitmap = JpegFiles.Load(uprightPath))
            {
                src = PixelBuffer.FromBitmap(bitmap);
            }
            var corners = new PointD[4];
            for (var i = 0; i < 4; i++)
            {
                corners[i] = new PointD(plan.Corners[i].X * src.Width, plan.Corners[i].Y * src.Height);
            }
            var page = PerspectiveWarp.Warp(src, corners);
            src = null;
            PageLook.Apply(page, PageLook.For(plan.Strength), color == PageColor.Grayscale);
            using (var output = page.ToBitmap())
            {
                JpegFiles.Save(output, outputPath, JpegFiles.PageQuality);
            }
        }

        public static void WriteRotated(string sourceJpeg, int quarterTurns, string outputPath)
        {
            using (var bitmap = JpegFiles.Load(sourceJpeg))
            {
                bitmap.RotateFlip(QuarterTurns(quarterTurns));
                JpegFiles.Save(bitmap, outputPath, JpegFiles.PageQuality);
            }
        }

        internal static RotateFlipType QuarterTurns(int turns)
        {
            switch (((turns % 4) + 4) % 4)
            {
                case 1: return RotateFlipType.Rotate90FlipNone;
                case 2: return RotateFlipType.Rotate180FlipNone;
                case 3: return RotateFlipType.Rotate270FlipNone;
                default: return RotateFlipType.RotateNoneFlipNone;
            }
        }
    }
}
