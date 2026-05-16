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
        public async Task Execute_ReturnsErrorWhenQueryMissing()
        {
            var surface = new Surface { OnCount = _ => 0 };
            var tool = new OutlookCountMessagesTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public Func<SearchMessagesArgs, int> OnCount { get; set; }
            public override int CountMessages(SearchMessagesArgs args) => OnCount(args);
        }
    }
}
