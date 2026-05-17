using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookGetCurrentSelectionToolTests
    {
        [Fact]
        public async Task Execute_DefaultArgs_PassesIncludeFullBodiesFalseAndMax5()
        {
            bool capturedIncludeBodies = true;     // start true so we detect "default to false"
            int capturedMaxItems = -1;
            var surface = new Surface
            {
                OnGetSelection = (incl, max) =>
                {
                    capturedIncludeBodies = incl;
                    capturedMaxItems = max;
                    return new CurrentSelectionResult
                    {
                        Folder = "Inbox",
                        FolderId = "fld_root",
                        Count = 0,
                        Messages = new MessageDetail[0],
                    };
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();

            var resultJson = await tool.ExecuteAsync("{}", surface, CancellationToken.None);

            Assert.False(capturedIncludeBodies);
            Assert.Equal(5, capturedMaxItems);
            var result = JObject.Parse(resultJson);
            Assert.Equal("Inbox", (string)result["folder"]);
            Assert.Equal(0, (int)result["count"]);
            Assert.NotNull(result["messages"]);
            Assert.Empty((JArray)result["messages"]);
        }

        [Fact]
        public async Task Execute_RespectsMaxItems_PassedToSurface()
        {
            int capturedMaxItems = -1;
            var surface = new Surface
            {
                OnGetSelection = (incl, max) =>
                {
                    capturedMaxItems = max;
                    return new CurrentSelectionResult { Messages = new MessageDetail[0] };
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();

            await tool.ExecuteAsync("{\"max_items\":3,\"include_full_bodies\":true}", surface, CancellationToken.None);

            Assert.Equal(3, capturedMaxItems);
        }

        [Fact]
        public async Task Execute_ClampsMaxItemsToHardCap()
        {
            int capturedMaxItems = -1;
            var surface = new Surface
            {
                OnGetSelection = (incl, max) =>
                {
                    capturedMaxItems = max;
                    return new CurrentSelectionResult { Messages = new MessageDetail[0] };
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();

            await tool.ExecuteAsync("{\"max_items\":999}", surface, CancellationToken.None);

            Assert.Equal(20, capturedMaxItems);   // hard cap per the spec
        }

        [Fact]
        public async Task Execute_ProjectsMessageDetail_ToWireShape()
        {
            var msg = new MessageDetail
            {
                Id = "msg_abc",
                Subject = "Re: Q4 plan",
                From = "Jane Doe <jane@acme.com>",
                ReceivedAt = new DateTimeOffset(2026, 5, 17, 9, 14, 0, TimeSpan.Zero),
                BodyPlaintext = "Hi team --- thoughts on regional split...",
                BodyTruncated = false,
                Attachments = new[] { new AttachmentSummary { Filename = "plan.xlsx", SizeBytes = 4096 } },
                ConversationTopic = "Q4 plan",
            };
            var surface = new Surface
            {
                OnGetSelection = (incl, max) => new CurrentSelectionResult
                {
                    Folder = "Inbox",
                    FolderId = "fld_inbox",
                    Count = 1,
                    Messages = new[] { msg },
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();

            var resultJson = await tool.ExecuteAsync("{}", surface, CancellationToken.None);

            var result = JObject.Parse(resultJson);
            Assert.Equal("Inbox", (string)result["folder"]);
            Assert.Equal("fld_inbox", (string)result["folder_id"]);
            Assert.Equal(1, (int)result["count"]);
            var arr = (JArray)result["messages"];
            Assert.Single(arr);
            var m = (JObject)arr[0];
            Assert.Equal("msg_abc", (string)m["id"]);
            Assert.Equal("Re: Q4 plan", (string)m["subject"]);
            Assert.Equal("Jane Doe <jane@acme.com>", (string)m["from"]);
            Assert.Equal("Q4 plan", (string)m["conversation_topic"]);
            Assert.True((bool)m["has_attachments"]);
            // Default include_full_bodies=false -> snippet present, no body.
            Assert.NotNull(m["snippet"]);
            Assert.Null(m["body_plaintext"]);
        }

        [Fact]
        public async Task Execute_IncludeFullBodiesTrue_AddsBodyPlaintext()
        {
            var msg = new MessageDetail
            {
                Id = "msg_1",
                BodyPlaintext = "full body here",
                BodyTruncated = false,
            };
            var surface = new Surface
            {
                OnGetSelection = (incl, max) => new CurrentSelectionResult
                {
                    Folder = "Inbox",
                    Count = 1,
                    Messages = new[] { msg },
                }
            };
            var tool = new OutlookGetCurrentSelectionTool();

            var resultJson = await tool.ExecuteAsync("{\"include_full_bodies\":true}", surface, CancellationToken.None);

            var m = (JObject)((JArray)JObject.Parse(resultJson)["messages"])[0];
            Assert.Equal("full body here", (string)m["body_plaintext"]);
            Assert.False((bool)m["body_truncated"]);
            Assert.Null(m["snippet"]);    // not emitted when bodies are included
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<bool, int, CurrentSelectionResult> OnGetSelection { get; set; }
            public override CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems)
                => OnGetSelection(includeFullBodies, maxItems);
        }
    }
}
