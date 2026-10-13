using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PagePlannerTests
    {
        private const string Plan =
            "{\"kind\":\"document_page\",\"reason\":\"page\",\"corners\":[[0.1,0.1],[0.9,0.1],[0.9,0.9],[0.1,0.9]],\"color\":\"grayscale\",\"strength\":\"normal\"}";

        private sealed class FakeAi : IStructuredCompletion
        {
            public readonly Queue<Func<CancellationToken, Task<string>>> Answers = new Queue<Func<CancellationToken, Task<string>>>();
            public readonly List<IReadOnlyList<JObject>> Parts = new List<IReadOnlyList<JObject>>();
            public string Effort;
            public string SchemaName;
            public int Calls;

            public Task<string> CompleteStructuredAsync(string instructions, IReadOnlyList<JObject> contentParts, string schemaName, JObject schema, string effort, CancellationToken cancellationToken)
            {
                Interlocked.Increment(ref Calls);
                lock (Parts) Parts.Add(contentParts);
                Effort = effort;
                SchemaName = schemaName;
                Func<CancellationToken, Task<string>> answer;
                lock (Answers) answer = Answers.Dequeue();
                return answer(cancellationToken);
            }
        }

        private static Func<CancellationToken, Task<string>> Returns(string text) => _ => Task.FromResult(text);

        private static Func<CancellationToken, Task<string>> Throws(Exception ex) => _ =>
        {
            var tcs = new TaskCompletionSource<string>();
            tcs.SetException(ex);
            return tcs.Task;
        };

        [Fact]
        public async Task PlanAsync_SendsTheImageAsADataUrl_AndParsesThePlan()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Returns(Plan));

            var plan = await new PagePlanner(ai).PlanAsync(new byte[] { 1, 2, 3 }, CancellationToken.None);

            Assert.True(plan.Analyzed);
            Assert.Equal("data:image/jpeg;base64,AQID", (string)ai.Parts[0][1]["image_url"]);
            Assert.Equal("input_text", (string)ai.Parts[0][0]["type"]);
            Assert.Equal("Low", ai.Effort);
            Assert.Equal(PagePlan.SchemaName, ai.SchemaName);
        }

        [Fact]
        public async Task PlanAsync_RetriesOnceAfterARateLimit_WaitingRetryAfter()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new CodexBackendException((HttpStatusCode)429, TimeSpan.FromSeconds(7), "slow down")));
            ai.Answers.Enqueue(Returns(Plan));
            var waited = new List<TimeSpan>();

            var plan = await new PagePlanner(ai, (d, ct) => { waited.Add(d); return Task.CompletedTask; })
                .PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.True(plan.Analyzed);
            Assert.Equal(new[] { TimeSpan.FromSeconds(7) }, waited);
            Assert.Equal(2, ai.Calls);
        }

        [Fact]
        public async Task PlanAsync_SecondRateLimit_FallsBack()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new CodexBackendException((HttpStatusCode)429, null, "slow down")));
            ai.Answers.Enqueue(Throws(new CodexBackendException((HttpStatusCode)429, null, "slow down")));

            var plan = await new PagePlanner(ai, (d, ct) => Task.CompletedTask).PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.False(plan.Analyzed);
            Assert.Equal(2, ai.Calls);
        }

        [Fact]
        public void RetryDelay_IsCappedAndDefaulted()
        {
            Assert.Equal(TimeSpan.FromSeconds(5), PagePlanner.RetryDelay(null));
            Assert.Equal(TimeSpan.FromSeconds(20), PagePlanner.RetryDelay(TimeSpan.FromSeconds(60)));
            Assert.Equal(TimeSpan.Zero, PagePlanner.RetryDelay(TimeSpan.FromSeconds(-3)));
        }

        [Fact]
        public async Task PlanAsync_BackendError_FallsBack()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new CodexBackendException(HttpStatusCode.InternalServerError, null, "boom")));

            var plan = await new PagePlanner(ai).PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.False(plan.Analyzed);
        }

        [Fact]
        public async Task PlanAsync_Timeout_FallsBack()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(async ct => { await Task.Delay(Timeout.Infinite, ct); return Plan; });

            var plan = await new PagePlanner(ai, timeout: TimeSpan.FromMilliseconds(50)).PlanAsync(new byte[] { 1 }, CancellationToken.None);

            Assert.False(plan.Analyzed);
        }

        [Fact]
        public async Task PlanAsync_NotSignedIn_Throws()
        {
            var ai = new FakeAi();
            ai.Answers.Enqueue(Throws(new NotSignedInException("not signed in", "Not signed in")));

            await Assert.ThrowsAsync<NotSignedInException>(() => new PagePlanner(ai).PlanAsync(new byte[] { 1 }, CancellationToken.None));
        }

        [Fact]
        public async Task PlanAsync_Cancelled_MakesNoCall()
        {
            var ai = new FakeAi();
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();

                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new PagePlanner(ai).PlanAsync(new byte[] { 1 }, cts.Token));
            }
            Assert.Equal(0, ai.Calls);
        }

        [Fact]
        public async Task PlanAsync_RunsAtMostThreeAtOnce()
        {
            var ai = new FakeAi();
            var release = new TaskCompletionSource<string>();
            var running = 0;
            var peak = 0;
            for (var i = 0; i < 5; i++)
            {
                ai.Answers.Enqueue(async ct =>
                {
                    var now = Interlocked.Increment(ref running);
                    lock (ai) peak = Math.Max(peak, now);
                    var text = await release.Task;
                    Interlocked.Decrement(ref running);
                    return text;
                });
            }
            var planner = new PagePlanner(ai);

            var tasks = Enumerable.Range(0, 5).Select(_ => Task.Run(() => planner.PlanAsync(new byte[] { 1 }, CancellationToken.None))).ToList();
            await Task.Delay(300);
            release.SetResult(Plan);
            await Task.WhenAll(tasks);

            Assert.Equal(PagePlanner.MaxConcurrent, peak);
        }

        [Fact]
        public void Instructions_CountIdsAsDocuments_AndNeverAlterContent()
        {
            Assert.Contains("ID or passport", PagePlanner.Instructions);
            Assert.Contains("never alter", PagePlanner.Instructions);
            Assert.Contains("top-left, top-right, bottom-right, bottom-left", PagePlanner.Instructions);
        }
    }
}
