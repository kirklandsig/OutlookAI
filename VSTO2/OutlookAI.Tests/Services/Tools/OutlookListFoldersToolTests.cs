using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookListFoldersToolTests
    {
        [Fact]
        public async Task Execute_ProjectsFolderListToJson()
        {
            var surface = new Surface
            {
                Folders = new List<FolderResult>
                {
                    new FolderResult { Id = "f1", Name = "Inbox", ParentId = null, ItemCount = 42 },
                    new FolderResult { Id = "f2", Name = "Sent", ParentId = "f1", ItemCount = 10 },
                }
            };
            var tool = new OutlookListFoldersTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"id\":\"f1\"", json);
            Assert.Contains("\"name\":\"Inbox\"", json);
            Assert.Contains("\"parent_id\":null", json);
            Assert.Contains("\"item_count\":42", json);
            Assert.Contains("\"parent_id\":\"f1\"", json);
        }

        [Fact]
        public async Task Execute_EmptyFoldersStillProducesArray()
        {
            var surface = new Surface { Folders = new List<FolderResult>() };
            var tool = new OutlookListFoldersTool();
            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);
            Assert.Contains("\"folders\":[]", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public List<FolderResult> Folders { get; set; }
            public override IReadOnlyList<FolderResult> ListFolders() => Folders;
        }
    }
}
