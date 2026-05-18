using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookCountMessagesToolTests
    {
        [Fact]
        public async Task Execute_ProjectsCountAndPassesArgsThrough()
        {
            SearchMessagesArgs observed = null;
            var surface = new Surface
            {
                OnCount = a => { observed = a; return 17; }
            };
            var tool = new OutlookCountMessagesTool();
            var json = await tool.ExecuteAsync(
                "{\"query\":\"newsletter\",\"folder_id\":\"f1\"}",
                surface, CancellationToken.None);
            Assert.Contains("\"count\":17", json);
            Assert.Equal("newsletter", observed.Query);
            Assert.Equal("f1", observed.FolderId);
        }

        [Fact]
        public async Task Execute_EmptyArgs_DefaultsToInboxNoFilter()
        {
            // Phase 3a removed the "query required" constraint.
            SearchMessagesArgs observed = null;
            var surface = new Surface { OnCount = a => { observed = a; return 0; } };
            var tool = new OutlookCountMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);

            Assert.DoesNotContain("\"error\":", json);
            Assert.Contains("\"count\":0", json);
            Assert.NotNull(observed);
            Assert.Null(observed.Query);
        }

        [Fact]
        public async Task Execute_ParsesAllStructuredFields_AndPassesToSurface()
        {
            SearchMessagesArgs observed = null;
            var surface = new Surface
            {
                OnCount = a => { observed = a; return 7; }
            };
            var tool = new OutlookCountMessagesTool();
            var argsJson = "{"
                + "\"from\":\"jane@acme.com\","
                + "\"is_unread\":true,"
                + "\"has_attachment\":true,"
                + "\"importance\":\"high\"}";

            var json = await tool.ExecuteAsync(argsJson, surface, CancellationToken.None);

            Assert.NotNull(observed);
            Assert.Equal("jane@acme.com", observed.From);
            Assert.Equal(true, observed.IsUnread);
            Assert.Equal(true, observed.HasAttachment);
            Assert.Equal("high", observed.Importance);
            Assert.Contains("\"count\":7", json);
        }

        [Fact]
        public async Task Execute_UsesSharedParser_ForScopeAndTriStates()
        {
            SearchMessagesArgs observed = null;
            var surface = new Surface { OnCount = a => { observed = a; return 5; } };
            var tool = new OutlookCountMessagesTool();

            await tool.ExecuteAsync(
                "{\"scope\":\"all_mail\",\"read_status\":\"read\"}",
                surface, CancellationToken.None);

            Assert.Equal("all_mail", observed.Scope);
            Assert.Equal("read", observed.ReadStatus);
            Assert.Equal(int.MaxValue, observed.MaxResults);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<SearchMessagesArgs, int> OnCount { get; set; }
            public override int CountMessages(SearchMessagesArgs args) => OnCount(args);
        }
    }
}
