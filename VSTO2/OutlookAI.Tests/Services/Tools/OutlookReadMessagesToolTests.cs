using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookReadMessagesToolTests
    {
        [Fact]
        public async Task Execute_PassesIdsThrough()
        {
            string[] observed = null;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = ids;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\",\"b\",\"c\"]}", surface, CancellationToken.None);
            Assert.Equal(new[] { "a", "b", "c" }, observed);
        }

        [Fact]
        public async Task Execute_DefaultIncludeBody_IsTrue()
        {
            bool? observed = null;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = includeBody;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, CancellationToken.None);
            Assert.True(observed);
        }

        [Fact]
        public async Task Execute_DefaultMaxItems_Is25()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = maxItems;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, CancellationToken.None);
            Assert.Equal(25, observed);
        }

        [Fact]
        public async Task Execute_MaxItemsClampedTo100()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = maxItems;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"],\"max_items\":9999}", surface, CancellationToken.None);
            Assert.Equal(100, observed);
        }

        [Fact]
        public async Task Execute_IncludeBodyFalse_Honored()
        {
            bool? observed = null;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) =>
                {
                    observed = includeBody;
                    return new MessageDetail[0];
                }
            };
            var tool = new OutlookReadMessagesTool();
            await tool.ExecuteAsync("{\"ids\":[\"a\"],\"include_body\":false}", surface, CancellationToken.None);
            Assert.False(observed);
        }

        [Fact]
        public async Task Execute_ProjectsMessageDetailsToJson()
        {
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => new[]
                {
                    new MessageDetail
                    {
                        Id = "m1",
                        Subject = "Q4 plan",
                        From = "jane@example.com",
                        To = new[] { "bob@example.com" },
                        Cc = new string[0],
                        ReceivedAt = DateTimeOffset.Parse("2026-05-14T18:32:00Z"),
                        BodyPlaintext = "Body text",
                        BodyTruncated = false,
                        Attachments = new AttachmentSummary[0],
                        InReplyToMessageId = null,
                        ConversationTopic = "Q4",
                    }
                }
            };
            var tool = new OutlookReadMessagesTool();
            var json = await tool.ExecuteAsync("{\"ids\":[\"m1\"]}", surface, CancellationToken.None);
            Assert.Contains("\"id\":\"m1\"", json);
            Assert.Contains("\"subject\":\"Q4 plan\"", json);
            Assert.Contains("\"body_plaintext\":\"Body text\"", json);
            Assert.Contains("\"body_truncated\":false", json);
            Assert.Contains("\"conversation_topic\":\"Q4\"", json);
        }

        [Fact]
        public async Task Execute_EmptyIds_ReturnsEmptyArrayWithoutCallingSurface()
        {
            bool called = false;
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => { called = true; return new MessageDetail[0]; }
            };
            var tool = new OutlookReadMessagesTool();
            var json = await tool.ExecuteAsync("{\"ids\":[]}", surface, CancellationToken.None);
            Assert.False(called);
            Assert.Contains("\"messages\":[]", json);
        }

        [Fact]
        public async Task Execute_MissingIds_ReturnsEmptyArray()
        {
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => new MessageDetail[0]
            };
            var tool = new OutlookReadMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"messages\":[]", json);
        }

        [Fact]
        public async Task Execute_PassesCancellationTokenThroughToSurface()
        {
            CancellationToken observed = default(CancellationToken);
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => { observed = ct; return new MessageDetail[0]; }
            };
            var tool = new OutlookReadMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, cts.Token);
                Assert.Equal(cts.Token, observed);
            }
        }

        [Fact]
        public async Task Execute_OnSurfaceCancellation_EmitsStructuredCancelEnvelope()
        {
            var surface = new Surface
            {
                OnReadMessages = (ids, includeBody, maxItems, ct) => throw new OperationCanceledException(ct)
            };
            var tool = new OutlookReadMessagesTool();
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();
                var json = await tool.ExecuteAsync("{\"ids\":[\"a\"]}", surface, cts.Token);
                Assert.Contains("\"error\"", json);
                Assert.Contains("\"code\":\"cancelled\"", json);
            }
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<string[], bool, int, CancellationToken, IReadOnlyList<MessageDetail>> OnReadMessages { get; set; }
            public override IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken))
                => OnReadMessages(ids, includeBody, maxItems, ct);
        }
    }
}
