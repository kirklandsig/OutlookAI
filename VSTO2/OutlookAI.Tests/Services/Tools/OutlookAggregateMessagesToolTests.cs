using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookAggregateMessagesToolTests
    {
        [Fact]
        public async Task Execute_ParsesAllFieldsAndPassesThrough()
        {
            AggregateMessagesArgs observed = null;
            var surface = new Surface
            {
                OnAggregate = (args, ct) => { observed = args; return new AggregationBucket[0]; }
            };
            var tool = new OutlookAggregateMessagesTool();
            var argsJson = "{"
                + "\"scope\":\"all_mail\","
                + "\"date_from\":\"2026-05-01T00:00:00Z\","
                + "\"date_to\":\"2026-05-31T00:00:00Z\","
                + "\"from\":\"jane@example.com\","
                + "\"group_by\":\"sender\","
                + "\"top_n\":25}";
            await tool.ExecuteAsync(argsJson, surface, CancellationToken.None);

            Assert.NotNull(observed);
            Assert.Equal("all_mail", observed.Scope);
            Assert.Equal(new DateTimeOffset(2026, 5, 1, 0, 0, 0, TimeSpan.Zero), observed.DateFrom);
            Assert.Equal(new DateTimeOffset(2026, 5, 31, 0, 0, 0, TimeSpan.Zero), observed.DateTo);
            Assert.Equal("jane@example.com", observed.From);
            Assert.Equal("sender", observed.GroupBy);
            Assert.Equal(25, observed.TopN);
        }

        [Fact]
        public async Task Execute_ProjectsBucketsAndTotal()
        {
            var surface = new Surface
            {
                OnAggregate = (args, ct) => new[]
                {
                    new AggregationBucket { Label = "Jane Doe", Count = 47 },
                    new AggregationBucket { Label = "Bob Smith", Count = 31 },
                }
            };
            var tool = new OutlookAggregateMessagesTool();
            var json = await tool.ExecuteAsync(
                "{\"scope\":\"all_mail\",\"group_by\":\"sender\"}",
                surface, CancellationToken.None);

            Assert.Contains("\"buckets\":[", json);
            Assert.Contains("\"label\":\"Jane Doe\"", json);
            Assert.Contains("\"count\":47", json);
            Assert.Contains("\"label\":\"Bob Smith\"", json);
            Assert.Contains("\"count\":31", json);
            Assert.Contains("\"total\":78", json);
        }

        [Fact]
        public async Task Execute_EmptyResult_ReturnsZeroTotal()
        {
            var surface = new Surface { OnAggregate = (args, ct) => new AggregationBucket[0] };
            var tool = new OutlookAggregateMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"buckets\":[]", json);
            Assert.Contains("\"total\":0", json);
        }

        [Fact]
        public async Task Execute_PassesCancellationTokenThroughToSurface()
        {
            CancellationToken observed = default(CancellationToken);
            var surface = new Surface
            {
                OnAggregate = (args, ct) => { observed = ct; return new AggregationBucket[0]; }
            };
            var tool = new OutlookAggregateMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                await tool.ExecuteAsync("{}", surface, cts.Token);
                Assert.Equal(cts.Token, observed);
            }
        }

        [Fact]
        public async Task Execute_OnSurfaceCancellation_EmitsStructuredCancelEnvelope()
        {
            var surface = new Surface
            {
                OnAggregate = (args, ct) => throw new OperationCanceledException(ct)
            };
            var tool = new OutlookAggregateMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();
                var json = await tool.ExecuteAsync("{}", surface, cts.Token);
                Assert.Contains("\"error\"", json);
                Assert.Contains("\"code\":\"cancelled\"", json);
            }
        }

        [Fact]
        public async Task Execute_TopNDefault10_PassedThrough()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnAggregate = (args, ct) => { observed = args.TopN; return new AggregationBucket[0]; }
            };
            var tool = new OutlookAggregateMessagesTool();
            await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Equal(10, observed);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<AggregateMessagesArgs, CancellationToken, IReadOnlyList<AggregationBucket>> OnAggregate { get; set; }
            public override IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken))
                => OnAggregate(args, ct);
        }
    }
}
