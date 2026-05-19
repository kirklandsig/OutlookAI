using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookSetCategoryToolTests
    {
        [Fact]
        public async Task Execute_CallsSurface_AndReturnsOk()
        {
            string observedId = null, observedCat = null;
            var surface = new Surface { OnSet = (i, c) => { observedId = i; observedCat = c; } };
            var tool = new OutlookSetCategoryTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"m1\",\"category\":\"Important\"}",
                surface, CancellationToken.None);
            Assert.Contains("\"ok\":true", json);
            Assert.Equal("m1", observedId);
            Assert.Equal("Important", observedCat);
        }

        [Fact]
        public async Task Execute_RejectsEmptyCategory()
        {
            var surface = new Surface { OnSet = (_, __) => { } };
            var tool = new OutlookSetCategoryTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"m1\",\"category\":\"\"}",
                surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public Action<string, string> OnSet { get; set; }
            public override void SetCategory(string messageId, string category) => OnSet(messageId, category);
        }
    }
}
