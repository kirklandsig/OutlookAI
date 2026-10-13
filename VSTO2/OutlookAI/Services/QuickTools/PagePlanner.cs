using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.QuickTools
{
    /// <summary>Asks the AI for each image's page plan, 3 at a time.</summary>
    internal sealed class PagePlanner
    {
        public const int MaxConcurrent = 3;
        public const int AiMaxEdge = 1600;
        public static readonly TimeSpan PerImageTimeout = TimeSpan.FromSeconds(90);
        private static readonly TimeSpan DefaultRetryDelay = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan MaxRetryDelay = TimeSpan.FromSeconds(20);

        internal const string Instructions =
            "You plan the cleanup of a photographed or scanned document for a program that applies your answer. "
            + "You never alter or retype content: the program only straightens, crops and adjusts exposure.\n"
            + "Decide whether the image is a document: a photo or scan of a page, card, ID or passport, form, receipt or screen, "
            + "even when it shows a person. Logos, signature images, decorations and photos with no document in them are not documents.\n"
            + "For a document, give the four corners of the document itself (not the background) as fractions of the image width "
            + "and height, in the order top-left, top-right, bottom-right, bottom-left as the text reads, so a sideways or "
            + "upside-down photo comes out upright. If the document fills the image or its edges aren't visible, give the image "
            + "corners in reading order. For anything else, give [[0,0],[1,0],[1,1],[0,1]].\n"
            + "Choose \"grayscale\" unless color carries meaning (stamps, highlights, colored ink, photos printed on the page), then \"color\".\n"
            + "Choose the strength from how dim, tinted or unevenly lit the image is: \"light\" for a clean scan, \"normal\" for a "
            + "typical phone photo, \"strong\" for a dark or shadowed one. Choose \"light\" when the document has large dark or "
            + "colored areas (a photo on an ID card, black bars), which a stronger cleanup would wash out.\n"
            + "Keep \"reason\" short: what the image is, e.g. \"lease page 2\" or \"company logo\".";

        internal const string Request = "Plan the cleanup of this image.";

        private readonly IStructuredCompletion _ai;
        private readonly Func<TimeSpan, CancellationToken, Task> _delay;
        private readonly TimeSpan _timeout;
        private readonly SemaphoreSlim _gate = new SemaphoreSlim(MaxConcurrent);

        public PagePlanner(IStructuredCompletion ai, Func<TimeSpan, CancellationToken, Task> delay = null, TimeSpan? timeout = null)
        {
            _ai = ai ?? throw new ArgumentNullException(nameof(ai));
            _delay = delay ?? Task.Delay;
            _timeout = timeout ?? PerImageTimeout;
        }

        /// <summary>The copy sent to the AI: at most 1600 px on the long edge, JPEG 85.</summary>
        public static byte[] PrepareForAi(Image page)
        {
            using (var small = JpegFiles.Scaled(page, AiMaxEdge))
            {
                return JpegFiles.Encode(small, JpegFiles.PageQuality);
            }
        }

        /// <summary>
        /// The image's plan, or the fallback plan when the AI fails, times out or
        /// answers implausibly. Throws only when the run is cancelled or nobody is
        /// signed in.
        /// </summary>
        public async Task<PagePlan> PlanAsync(byte[] jpeg, CancellationToken cancellationToken)
        {
            var parts = new List<JObject>
            {
                new JObject(new JProperty("type", "input_text"), new JProperty("text", Request)),
                new JObject(
                    new JProperty("type", "input_image"),
                    new JProperty("image_url", "data:image/jpeg;base64," + Convert.ToBase64String(jpeg))),
            };

            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                for (var attempt = 0; ; attempt++)
                {
                    try
                    {
                        using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                        {
                            timeout.CancelAfter(_timeout);
                            var text = await _ai.CompleteStructuredAsync(
                                Instructions, parts, PagePlan.SchemaName, PagePlan.Schema(), "Low", timeout.Token).ConfigureAwait(false);
                            return PagePlan.FromJson(text);
                        }
                    }
                    catch (CodexBackendException ex) when ((int)ex.StatusCode == 429 && attempt == 0)
                    {
                        await _delay(RetryDelay(ex.RetryAfter), cancellationToken).ConfigureAwait(false);
                    }
                }
            }
            catch (NotSignedInException)
            {
                throw;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                TraceLog.Write("QuickTools: planning failed (" + ex.GetType().Name + ")", "QuickTools");
                return PagePlan.Fallback();
            }
            finally
            {
                _gate.Release();
            }
        }

        internal static TimeSpan RetryDelay(TimeSpan? retryAfter)
        {
            if (!retryAfter.HasValue) return DefaultRetryDelay;
            if (retryAfter.Value < TimeSpan.Zero) return TimeSpan.Zero;
            return retryAfter.Value > MaxRetryDelay ? MaxRetryDelay : retryAfter.Value;
        }
    }
}
