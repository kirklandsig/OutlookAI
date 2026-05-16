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
        public async Task Execute_ReturnsStructuredErrorWhenQueryMissing()
        {
            var surface = new Surface { OnSearch = _ => new MessageSummary[0] };
            var tool = new OutlookSearchMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<SearchMessagesArgs, IReadOnlyList<MessageSummary>> OnSearch { get; set; }
            public override IReadOnlyList<MessageSummary> SearchMessages(SearchMessagesArgs args)
                => OnSearch(args);
        }
    }
}
