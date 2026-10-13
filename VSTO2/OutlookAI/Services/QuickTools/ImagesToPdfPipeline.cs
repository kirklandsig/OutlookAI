using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.QuickTools
{
    internal sealed class PipelineResult
    {
        public PipelineResult(IReadOnlyList<ReviewPage> pages, int skippedSmall, int notProcessed)
        {
            Pages = pages;
            SkippedSmall = skippedSmall;
            NotProcessed = notProcessed;
        }

        public IReadOnlyList<ReviewPage> Pages { get; }
        public int SkippedSmall { get; }
        public int NotProcessed { get; }

        /// <summary>Every readable page fell back (AI unreachable or every answer unusable).</summary>
        public bool NoneAnalyzed =>
            Pages.Any(p => p.Readable) && Pages.Where(p => p.Readable).All(p => !p.Plan.Analyzed);
    }

    /// <summary>
    /// Decode -> plan (AI) -> clean, writing every image as files in the run folder.
    /// The caller must run it off the UI thread (phase 1 is synchronous until the first await); at most one full-size image per worker in memory.
    /// </summary>
    internal sealed class ImagesToPdfPipeline
    {
        public const int MinShortSide = 200;
        public const int MaxPages = 30;
        public const int MaxParallelCleaning = 2;

        private readonly Func<byte[], CancellationToken, Task<PagePlan>> _plan;
        private readonly string _runFolder;

        // Test seam: lets a test make a frame fail to decode.
        internal Func<ImagePages, int, Bitmap> FrameDecoder { get; set; } = (pages, index) => pages.Decode(index);

        // Test seam: lets a test make opening an image fail.
        internal Func<string, ImagePages> Opener { get; set; } = path => ImageLoader.Open(path);

        public ImagesToPdfPipeline(Func<byte[], CancellationToken, Task<PagePlan>> plan, string runFolder)
        {
            _plan = plan ?? throw new ArgumentNullException(nameof(plan));
            _runFolder = runFolder ?? throw new ArgumentNullException(nameof(runFolder));
        }

        public async Task<PipelineResult> RunAsync(IReadOnlyList<CollectedImage> images, IProgress<string> progress, CancellationToken cancellationToken)
        {
            progress?.Report("Reading images\u2026");
            var pages = new List<ReviewPage>();
            var aiCopies = new Dictionary<ReviewPage, byte[]>();
            var skipped = 0;
            var notProcessed = 0;

            foreach (var image in images)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (image.Path == null)
                {
                    // Outlook couldn't save this attachment (see EmailImageCollector.SaveCopy).
                    pages.Add(new ReviewPage(pages.Count + 1, image.FileName));
                    continue;
                }
                ImagePages frames;
                try
                {
                    frames = Opener(image.Path);
                }
                catch (Exception ex) when (ex is UnreadableImageException || ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException || ex is OutOfMemoryException)
                {
                    TraceLog.Write("QuickTools: unreadable image (" + ex.GetType().Name + ")", "QuickTools");
                    pages.Add(new ReviewPage(pages.Count + 1, image.FileName));
                    continue;
                }

                using (frames)
                {
                    for (var f = 0; f < frames.Count; f++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var label = frames.Count > 1 ? image.FileName + " (page " + (f + 1) + ")" : image.FileName;
                        if (aiCopies.Count >= MaxPages)
                        {
                            // Over the cap: listed, never decoded (32-bit memory).
                            notProcessed++;
                            pages.Add(new ReviewPage(pages.Count + 1, label) { OverLimit = true });
                            continue;
                        }

                        Bitmap bitmap;
                        try
                        {
                            bitmap = FrameDecoder(frames, f);
                        }
                        catch (Exception ex) when (!(ex is OperationCanceledException))
                        {
                            TraceLog.Write("QuickTools: unreadable frame (" + ex.GetType().Name + ")", "QuickTools");
                            pages.Add(new ReviewPage(pages.Count + 1, label));
                            continue;
                        }

                        using (bitmap)
                        {
                            if (Math.Min(bitmap.Width, bitmap.Height) < MinShortSide)
                            {
                                skipped++;
                                continue;
                            }
                            var page = new ReviewPage(pages.Count + 1, label) { Readable = true };
                            page.UprightPath = Path.Combine(_runFolder, "upright-" + page.Number.ToString("D2", CultureInfo.InvariantCulture) + ".jpg");
                            JpegFiles.Save(bitmap, page.UprightPath, JpegFiles.OriginalQuality);
                            aiCopies[page] = PagePlanner.PrepareForAi(bitmap);
                            pages.Add(page);
                        }
                    }
                }
            }

            var readable = pages.Where(p => p.Readable).ToList();
            var done = 0;
            if (readable.Count > 0) progress?.Report("Analyzing image 1 of " + readable.Count + "\u2026");
            using (var cleaning = new SemaphoreSlim(MaxParallelCleaning))
            {
                var work = readable.Select(async page =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var plan = await _plan(aiCopies[page], cancellationToken).ConfigureAwait(false);
                    page.Plan = plan;
                    page.Include = plan.Kind == PageKind.DocumentPage;
                    await cleaning.WaitAsync(cancellationToken).ConfigureAwait(false);
                    try
                    {
                        page.CleanedPath = Path.Combine(_runFolder, "cleaned-" + page.Number.ToString("D2", CultureInfo.InvariantCulture) + ".jpg");
                        await Task.Run(() => PageCleaner.WriteCleaned(page.UprightPath, plan, page.EffectiveColor, page.CleanedPath), cancellationToken).ConfigureAwait(false);
                    }
                    catch (Exception ex) when (!(ex is OperationCanceledException))
                    {
                        TraceLog.Write("QuickTools: cleanup failed (" + ex.GetType().Name + ")", "QuickTools");
                        page.CleanupFailed = true;
                        page.UseOriginal = true;
                    }
                    finally
                    {
                        cleaning.Release();
                    }
                    var finished = Interlocked.Increment(ref done);
                    progress?.Report(finished < readable.Count
                        ? "Analyzing image " + (finished + 1) + " of " + readable.Count + "\u2026"
                        : "Finishing\u2026");
                }).ToList();
                await Task.WhenAll(work).ConfigureAwait(false);
            }

            return new PipelineResult(pages, skipped, notProcessed);
        }

        /// <summary>Re-cleans one page after the review changed its color.</summary>
        public static void Reclean(ReviewPage page)
        {
            PageCleaner.WriteCleaned(page.UprightPath, page.Plan, page.EffectiveColor, page.CleanedPath);
        }

        /// <summary>The JPEG files for the PDF, in order: rotated copies where the review turned a page.</summary>
        public static List<string> FinalJpegs(IEnumerable<ReviewPage> chosenInOrder, string runFolder)
        {
            var files = new List<string>();
            foreach (var page in chosenInOrder)
            {
                var source = page.UseOriginal ? page.UprightPath : page.CleanedPath;
                if (page.QuarterTurns % 4 == 0)
                {
                    files.Add(source);
                    continue;
                }
                var rotated = Path.Combine(runFolder, "final-" + page.Number.ToString("D2", CultureInfo.InvariantCulture) + ".jpg");
                PageCleaner.WriteRotated(source, page.QuarterTurns, rotated);
                files.Add(rotated);
            }
            return files;
        }
    }
}
