using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookReadMessageToolTests
    {
        [Fact]
        public async Task Execute_ProjectsMessageDetailToJson()
        {
            var surface = new Surface
            {
                OnRead = (id, full) => new MessageDetail
                {
                    Id = id,
                    Subject = "Hello",
                    From = "alice@example.com",
                    To = new[] { "bob@example.com" },
                    Cc = new string[0],
                    ReceivedAt = DateTimeOffset.Parse("2026-05-10T12:00:00Z"),
                    BodyPlaintext = "world",
                    BodyTruncated = false,
                    Attachments = new AttachmentSummary[0],
                    InReplyToMessageId = null,
                    ConversationTopic = "Hello",
                }
            };
            var tool = new OutlookReadMessageTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"m1\"}", surface, CancellationToken.None);
            Assert.Contains("\"id\":\"m1\"", json);
            Assert.Contains("\"subject\":\"Hello\"", json);
            Assert.Contains("\"body_plaintext\":\"world\"", json);
            Assert.Contains("\"in_reply_to_message_id\":null", json);
        }

        [Fact]
        public async Task Execute_ReturnsErrorWhenMessageIdMissing()
        {
            var surface = new Surface { OnRead = (_, __) => null };
            var tool = new OutlookReadMessageTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        [Fact]
        public async Task Execute_ReturnsNotFoundWhenSurfaceReturnsNull()
        {
            var surface = new Surface { OnRead = (_, __) => null };
            var tool = new OutlookReadMessageTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"unknown\"}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"not_found\"", json);
        }

        [Fact]
        public async Task Execute_DefaultsIncludeFullBodyToTrue()
        {
            bool? observed = null;
            var surface = new Surface
            {
                OnRead = (id, full) => { observed = full; return new MessageDetail
                {
                    Id = id, BodyPlaintext = "", ConversationTopic = "",
                    To = new string[0], Cc = new string[0],
                    Attachments = new AttachmentSummary[0],
                }; }
            };
            var tool = new OutlookReadMessageTool();
            await tool.ExecuteAsync("{\"message_id\":\"x\"}", surface, CancellationToken.None);
            Assert.True(observed);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<string, bool, MessageDetail> OnRead { get; set; }
            public override MessageDetail ReadMessage(string messageId, bool includeFullBody)
                => OnRead(messageId, includeFullBody);
        }
    }
}
