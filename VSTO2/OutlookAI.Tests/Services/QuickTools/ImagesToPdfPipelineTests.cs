using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class ImagesToPdfPipelineTests : IDisposable
    {
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-pipeline", Path.GetRandomFileName());

        public ImagesToPdfPipelineTests()
        {
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private CollectedImage AttachedImage(string name, int width, int height)
        {
            var path = Path.Combine(_dir, name);
            using (var page = SyntheticPages.Page(width, height, SyntheticPages.DimPaper, SyntheticPages.DimInk))
            {
                JpegFiles.Save(page, path, 90);
            }
            return new CollectedImage(path, name, isInline: true);
        }

        private static string PlanJson(string kind, string reason = "") =>
            "{\"kind\":\"" + kind + "\",\"reason\":\"" + reason + "\",\"corners\":[[0,0],[1,0],[1,1],[0,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}";

        // Plans by the AI copy's width (concurrency makes call order nondeterministic).
        private static Func<byte[], CancellationToken, Task<PagePlan>> PlansByWidth(Func<int, string> json) => (jpeg, ct) =>
        {
            int w, h, c;
            JpegHeader.TryRead(jpeg, out w, out h, out c);
            return Task.FromResult(PagePlan.FromJson(json(w)));
        };

        [Fact]
        public async Task RunAsync_CleansEachPage_AndIncludesOnlyDocuments()
        {
            var images = new[] { AttachedImage("scan.jpg", 850, 1100), AttachedImage("logo.jpg", 400, 300) };
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => w == 400 ? PlanJson("not_document", "company logo") : PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(images, null, CancellationToken.None);

            Assert.Equal(2, result.Pages.Count);
            Assert.True(result.Pages[0].Include);
            Assert.Equal("page", result.Pages[0].Status);
            Assert.True(File.Exists(result.Pages[0].CleanedPath));
            Assert.True(File.Exists(result.Pages[0].UprightPath));
            Assert.False(result.Pages[1].Include);
            Assert.Equal("not a document: company logo", result.Pages[1].Status);
            Assert.False(result.NoneAnalyzed);
        }

        [Fact]
        public async Task RunAsync_SkipsSmallImages()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(new[] { AttachedImage("icon.jpg", 150, 150) }, null, CancellationToken.None);

            Assert.Empty(result.Pages);
            Assert.Equal(1, result.SkippedSmall);
        }

        [Fact]
        public async Task RunAsync_StopsAt30Pages()
        {
            var images = Enumerable.Range(1, 31).Select(i => AttachedImage("p" + i + ".jpg", 220, 220)).ToList();
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(images, null, CancellationToken.None);

            Assert.Equal(30, result.Pages.Count(p => p.Readable));
            Assert.Equal(1, result.NotProcessed);
            Assert.Equal("not processed (limit 30)", result.Pages.Last().Status);
        }

        [Fact]
        public async Task RunAsync_UnreadableFile_IsListed()
        {
            var bad = Path.Combine(_dir, "broken.jpg");
            File.WriteAllText(bad, "nope");
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(new[] { new CollectedImage(bad, "broken.jpg", false) }, null, CancellationToken.None);

            var page = Assert.Single(result.Pages);
            Assert.False(page.Readable);
            Assert.False(page.Include);
            Assert.Equal("can't read this file", page.Status);
        }

        // An attachment Outlook couldn't save arrives with no path: listed, never dropped.
        [Fact]
        public async Task RunAsync_ImageThatCouldntBeSaved_IsListedInItsPlace()
        {
            var images = new[] { AttachedImage("a.jpg", 300, 400), new CollectedImage(null, "blocked.jpg", false), AttachedImage("c.jpg", 300, 400) };
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(images, null, CancellationToken.None);

            Assert.Equal(3, result.Pages.Count);
            Assert.True(result.Pages[0].Readable);
            Assert.False(result.Pages[1].Readable);
            Assert.Equal("blocked.jpg", result.Pages[1].Label);
            Assert.Equal("can't read this file", result.Pages[1].Status);
            Assert.True(result.Pages[2].Readable);
        }

        [Fact]
        public async Task RunAsync_CleanupFailure_KeepsThePageWithItsOriginal()
        {
            var pipeline = new ImagesToPdfPipeline((jpeg, ct) =>
            {
                // The upright copy is no longer a JPEG when the cleaner reads it.
                File.WriteAllText(Path.Combine(_dir, "upright-01.jpg"), "not a jpeg");
                return Task.FromResult(PagePlan.FromJson(PlanJson("document_page")));
            }, _dir);

            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400) }, null, CancellationToken.None);

            var page = Assert.Single(result.Pages);
            Assert.True(page.Readable);
            Assert.True(page.CleanupFailed);
            Assert.True(page.UseOriginal);
            Assert.True(page.Include);
            Assert.Equal("cleanup failed, original only", page.Status);
        }

        [Fact]
        public async Task RunAsync_NoneAnalyzed_WhenEveryPlanFellBack()
        {
            var pipeline = new ImagesToPdfPipeline((jpeg, ct) => Task.FromResult(PagePlan.Fallback()), _dir);

            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400) }, null, CancellationToken.None);

            Assert.True(result.NoneAnalyzed);
            Assert.Equal("AI couldn't analyze this one", result.Pages[0].Status);
            Assert.True(result.Pages[0].Include);
        }

        [Fact]
        public async Task RunAsync_Cancelled_Throws()
        {
            using (var cts = new CancellationTokenSource())
            {
                var calls = 0;
                var pipeline = new ImagesToPdfPipeline((jpeg, ct) =>
                {
                    Interlocked.Increment(ref calls);
                    cts.Cancel();
                    ct.ThrowIfCancellationRequested();
                    return Task.FromResult(PagePlan.Fallback());
                }, _dir);

                await Assert.ThrowsAnyAsync<OperationCanceledException>(
                    () => pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400), AttachedImage("b.jpg", 300, 400) }, null, cts.Token));
                Assert.True(calls >= 1);
            }
        }

        [Fact]
        public async Task Reclean_InColor_IsNoLongerGray()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);
            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400) }, null, CancellationToken.None);
            var page = result.Pages[0];
            using (var stamped = SyntheticPages.Page(300, 400, SyntheticPages.DimPaper, SyntheticPages.DimInk, stamp: Color.FromArgb(200, 40, 40)))
            {
                JpegFiles.Save(stamped, page.UprightPath, 95);
            }

            page.ColorOverride = PageColor.Color;
            ImagesToPdfPipeline.Reclean(page);

            using (var cleaned = JpegFiles.Load(page.CleanedPath))
            {
                var px = PixelBuffer.FromBitmap(cleaned);
                var colored = 0;
                for (var i = 0; i < px.Data.Length; i += 3) if (px.Data[i + 2] - px.Data[i + 1] > 60) colored++;
                Assert.True(colored > 20);
            }
        }

        [Fact]
        public async Task FinalJpegs_RespectOrderRotationAndOriginals()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);
            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400), AttachedImage("b.jpg", 400, 300) }, null, CancellationToken.None);
            var first = result.Pages[0];
            var second = result.Pages[1];
            second.QuarterTurns = 1;
            first.UseOriginal = true;

            var files = ImagesToPdfPipeline.FinalJpegs(new[] { second, first }, _dir);

            Assert.Equal(2, files.Count);
            using (var rotated = JpegFiles.Load(files[0])) Assert.True(rotated.Height > rotated.Width);
            Assert.Equal(first.UprightPath, files[1]);
        }

        [Fact]
        public async Task RunAsync_MultiPageTiff_GivesOnePagePerFrame()
        {
            var path = Path.Combine(_dir, "fax.tif");
            var bitmaps = Enumerable.Range(0, 3).Select(_ => SyntheticPages.Page(300, 400, SyntheticPages.DimPaper, SyntheticPages.DimInk)).ToList();
            try
            {
                var tiff = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Tiff.Guid);
                using (var parameters = new EncoderParameters(1))
                {
                    parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                    bitmaps[0].Save(path, tiff, parameters);
                    parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                    for (var i = 1; i < 3; i++) bitmaps[0].SaveAdd(bitmaps[i], parameters);
                    parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                    bitmaps[0].SaveAdd(parameters);
                }
            }
            finally { foreach (var b in bitmaps) b.Dispose(); }
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);

            var result = await pipeline.RunAsync(new[] { new CollectedImage(path, "fax.tif", false) }, null, CancellationToken.None);

            Assert.Equal(3, result.Pages.Count(p => p.Readable));
            Assert.Equal(new[] { "fax.tif (page 1)", "fax.tif (page 2)", "fax.tif (page 3)" }, result.Pages.Select(p => p.Label).ToArray());
        }

        [Fact]
        public async Task RunAsync_FrameThatFailsToDecode_IsListed_AndNextImageStillRuns()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);
            var calls = 0;
            pipeline.FrameDecoder = (pages, index) =>
            {
                if (Interlocked.Increment(ref calls) == 1) throw new FormatException("corrupt frame");
                return pages.Decode(index);
            };

            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400), AttachedImage("b.jpg", 300, 400) }, null, CancellationToken.None);

            Assert.Equal(2, result.Pages.Count);
            Assert.False(result.Pages[0].Readable);
            Assert.Equal("can't read this file", result.Pages[0].Status);
            Assert.True(result.Pages[1].Readable);
            Assert.Equal("page", result.Pages[1].Status);
        }

        [Fact]
        public async Task RunAsync_ImageThatRunsOutOfMemoryOnOpen_IsListed_AndNextImageStillRuns()
        {
            var pipeline = new ImagesToPdfPipeline(PlansByWidth(w => PlanJson("document_page")), _dir);
            var calls = 0;
            pipeline.Opener = path =>
            {
                if (Interlocked.Increment(ref calls) == 1) throw new OutOfMemoryException();
                return ImageLoader.Open(path);
            };

            var result = await pipeline.RunAsync(new[] { AttachedImage("a.jpg", 300, 400), AttachedImage("b.jpg", 300, 400) }, null, CancellationToken.None);

            Assert.Equal(2, result.Pages.Count);
            Assert.False(result.Pages[0].Readable);
            Assert.Equal("can't read this file", result.Pages[0].Status);
            Assert.True(result.Pages[1].Readable);
        }

        [Fact]
        public void RunFolder_Delete_IgnoresFoldersOutsideRoot()
        {
            var saved = QuickToolsRunFolder.Root;
            QuickToolsRunFolder.Root = Path.Combine(_dir, "root");
            try
            {
                var outside = Path.Combine(_dir, "elsewhere");
                Directory.CreateDirectory(outside);

                QuickToolsRunFolder.Delete(outside);

                Assert.True(Directory.Exists(outside));
            }
            finally
            {
                QuickToolsRunFolder.Root = saved;
            }
        }
        [Fact]
        public void RunFolder_IsCreated_AndOldOnesArePurged()
        {
            var root = Path.Combine(_dir, "root");
            var saved = QuickToolsRunFolder.Root;
            QuickToolsRunFolder.Root = root;
            try
            {
                var old = QuickToolsRunFolder.Create(new DateTime(2026, 10, 1, 9, 0, 0, DateTimeKind.Utc));
                Directory.SetCreationTimeUtc(old, new DateTime(2026, 10, 1, 9, 0, 0, DateTimeKind.Utc));

                var current = QuickToolsRunFolder.Create(new DateTime(2026, 10, 13, 9, 0, 0, DateTimeKind.Utc));

                Assert.True(Directory.Exists(current));
                Assert.False(Directory.Exists(old));
                QuickToolsRunFolder.Delete(current);
                Assert.False(Directory.Exists(current));
                QuickToolsRunFolder.Delete(current); // already gone: no exception
            }
            finally
            {
                QuickToolsRunFolder.Root = saved;
            }
        }
    }
}
