using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookListRecentThreadsWithToolTests
    {
        [Fact]
        public async Task Execute_ProjectsThreadList()
        {
            string observedRecipient = null;
            int observedMax = -1;
            var surface = new Surface
            {
                OnList = (r, m) =>
                {
                    observedRecipient = r;
                    observedMax = m;
                    return new[]
                    {
                        new ThreadSummary
                        {
                            ThreadTopic = "Q4",
                            LastMessageAt = DateTimeOffset.Parse("2026-05-10T12:00:00Z"),
                            MessageCount = 3,
                            Snippet = "lgtm",
                            ThreadId = "t1",
                        }
                    };
                }
            };
            var tool = new OutlookListRecentThreadsWithTool();
            var json = await tool.ExecuteAsync(
                "{\"recipient_email\":\"alice@example.com\",\"max_threads\":3}",
                surface, CancellationToken.None);
            Assert.Contains("\"thread_topic\":\"Q4\"", json);
            Assert.Contains("\"thread_id\":\"t1\"", json);
            Assert.Equal("alice@example.com", observedRecipient);
            Assert.Equal(3, observedMax);
        }

        [Fact]
        public async Task Execute_ClampsMaxThreadsToHardCap()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnList = (r, m) => { observed = m; return new ThreadSummary[0]; }
            };
            var tool = new OutlookListRecentThreadsWithTool();
            await tool.ExecuteAsync(
                "{\"recipient_email\":\"a@b.com\",\"max_threads\":9999}",
                surface, CancellationToken.None);
            Assert.Equal(20, observed);
        }

        [Fact]
        public async Task Execute_ReturnsErrorWhenRecipientMissing()
        {
            var surface = new Surface { OnList = (_, __) => new ThreadSummary[0] };
            var tool = new OutlookListRecentThreadsWithTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<string, int, IReadOnlyList<ThreadSummary>> OnList { get; set; }
            public override IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads)
                => OnList(recipientEmail, maxThreads);
        }
    }
}
