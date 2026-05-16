using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookMarkAsReadToolTests
    {
        [Fact]
        public async Task Execute_CallsSurface_AndReturnsOk()
        {
            string observedId = null; bool? observedRead = null;
            var surface = new Surface
            {
                OnMark = (id, r) => { observedId = id; observedRead = r; }
            };
            var tool = new OutlookMarkAsReadTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"m1\",\"read\":true}", surface, CancellationToken.None);
            Assert.Contains("\"ok\":true", json);
            Assert.Equal("m1", observedId);
            Assert.True(observedRead);
        }

        [Fact]
        public async Task Execute_ReturnsErrorWhenArgsInvalid()
        {
            var surface = new Surface { OnMark = (_, __) => { } };
            var tool = new OutlookMarkAsReadTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"m1\"}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public Action<string, bool> OnMark { get; set; }
            public override void MarkAsRead(string messageId, bool read) => OnMark(messageId, read);
        }
    }
}
