using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookSearchMessagesToolTests
    {
        [Fact]
        public async Task Execute_ProjectsHitsAndPassesArgsThrough()
        {
            SearchMessagesArgs observed = null;
            var surface = new Surface
            {
                OnSearch = a => { observed = a; return new[]
                {
                    new MessageSummary
                    {
                        Id = "m1",
                        Subject = "Q4",
                        From = "alice@example.com",
                        To = new[] { "bob@example.com" },
                        ReceivedAt = DateTimeOffset.Parse("2026-05-10T12:00:00Z"),
                        Snippet = "ready?",
                        HasAttachments = true,
                    }
                }; }
            };
            var tool = new OutlookSearchMessagesTool();
            var json = await tool.ExecuteAsync(
                "{\"query\":\"Q4\",\"folder_id\":\"f1\",\"max_results\":50}",
                surface, CancellationToken.None);

            Assert.Contains("\"id\":\"m1\"", json);
            Assert.Contains("\"subject\":\"Q4\"", json);
            Assert.Contains("\"has_attachments\":true", json);
            Assert.NotNull(observed);
            Assert.Equal("Q4", observed.Query);
            Assert.Equal("f1", observed.FolderId);
            Assert.Equal(50, observed.MaxResults);
        }

        [Fact]
        public async Task Execute_ClampsMaxResultsToHardCap()
        {
            int observed = -1;
            var surface = new Surface
            {
                OnSearch = a => { observed = a.MaxResults; return new MessageSummary[0]; }
            };
            var tool = new OutlookSearchMessagesTool();
            await tool.ExecuteAsync(
                "{\"query\":\"q\",\"max_results\":9999}",
                surface, CancellationToken.None);
            Assert.Equal(100, observed);
        }

        [Fact]
        public async Task Execute_EmptyArgs_DefaultsToInboxNoFilter()
        {
            // Phase 3a removed the "query required" constraint. Zero-arg
            // calls now succeed; surface gets a SearchMessagesArgs with
            // all-null fields and MaxResults=25 (default).
            SearchMessagesArgs observed = null;
            var surface = new Surface
            {
                OnSearch = a => { observed = a; return new MessageSummary[0]; }
            };
            var tool = new OutlookSearchMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);

            Assert.DoesNotContain("\"error\":", json);
            Assert.NotNull(observed);
            Assert.Null(observed.Query);
            Assert.Null(observed.From);
            Assert.Null(observed.HasAttachment);
            Assert.Equal(25, observed.MaxResults);
        }

        [Fact]
        public async Task Execute_ParsesAllStructuredFields_AndPassesToSurface()
        {
            SearchMessagesArgs observed = null;
            var surface = new Surface
            {
                OnSearch = a => { observed = a; return new MessageSummary[0]; }
            };
            var tool = new OutlookSearchMessagesTool();
            var argsJson = "{"
                + "\"query\":\"Q4\","
                + "\"from\":\"jane@acme.com\","
                + "\"subject_contains\":\"plan\","
                + "\"body_contains\":\"draft\","
                + "\"has_attachment\":true,"
                + "\"is_unread\":true,"
                + "\"is_flagged\":false,"
                + "\"importance\":\"high\","
                + "\"date_from\":\"2026-05-10T00:00:00Z\","
                + "\"date_to\":\"2026-05-17T00:00:00Z\","
                + "\"max_results\":50}";

            await tool.ExecuteAsync(argsJson, surface, CancellationToken.None);

            Assert.NotNull(observed);
            Assert.Equal("Q4", observed.Query);
            Assert.Equal("jane@acme.com", observed.From);
            Assert.Equal("plan", observed.SubjectContains);
            Assert.Equal("draft", observed.BodyContains);
            Assert.Equal(true, observed.HasAttachment);
            Assert.Equal(true, observed.IsUnread);
            // Hidden old-shape false values are ignored so model defaults do
            // not pollute searches. Use flag_status=unflagged for an explicit
            // negative flag filter.
            Assert.Null(observed.IsFlagged);
            Assert.Equal("high", observed.Importance);
            Assert.Equal(50, observed.MaxResults);
            Assert.Equal(new DateTimeOffset(2026, 5, 10, 0, 0, 0, TimeSpan.Zero), observed.DateFrom);
            Assert.Equal(new DateTimeOffset(2026, 5, 17, 0, 0, 0, TimeSpan.Zero), observed.DateTo);
        }

        [Fact]
        public async Task Execute_UsesSharedParser_NewTriStateAndScopeFields()
        {
            SearchMessagesArgs observed = null;
            var surface = new Surface
            {
                OnSearch = a => { observed = a; return new MessageSummary[0]; }
            };
            var tool = new OutlookSearchMessagesTool();
            var argsJson = "{"
                + "\"scope\":\"all_mail\","
                + "\"sort_order\":\"oldest\","
                + "\"read_status\":\"unread\","
                + "\"attachment_filter\":\"with\","
                + "\"flag_status\":\"flagged\","
                + "\"importance_filter\":\"high\"}";

            await tool.ExecuteAsync(argsJson, surface, CancellationToken.None);

            Assert.Equal("all_mail", observed.Scope);
            Assert.Equal("oldest", observed.SortOrder);
            Assert.Equal("unread", observed.ReadStatus);
            Assert.Equal("with", observed.AttachmentFilter);
            Assert.Equal("flagged", observed.FlagStatus);
            Assert.Equal("high", observed.ImportanceFilter);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<SearchMessagesArgs, IReadOnlyList<MessageSummary>> OnSearch { get; set; }
            public override IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args)
                => OnSearch(args);
        }
    }
}
