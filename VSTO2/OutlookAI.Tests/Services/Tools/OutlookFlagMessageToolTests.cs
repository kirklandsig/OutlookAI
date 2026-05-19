using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookFlagMessageToolTests
    {
        [Theory]
        [InlineData("none")]
        [InlineData("todo")]
        [InlineData("complete")]
        public async Task Execute_AcceptsValidFlagValues(string flag)
        {
            string observedFlag = null;
            var surface = new Surface { OnFlag = (_, f) => observedFlag = f };
            var tool = new OutlookFlagMessageTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"m1\",\"flag\":\"" + flag + "\"}",
                surface, CancellationToken.None);
            Assert.Contains("\"ok\":true", json);
            Assert.Equal(flag, observedFlag);
        }

        [Fact]
        public async Task Execute_RejectsUnknownFlag()
        {
            var surface = new Surface { OnFlag = (_, __) => { } };
            var tool = new OutlookFlagMessageTool();
            var json = await tool.ExecuteAsync(
                "{\"message_id\":\"m1\",\"flag\":\"snooze\"}", surface, CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", json);
        }

        private sealed class Surface : MinimalSurface
        {
            public Action<string, string> OnFlag { get; set; }
            public override void FlagMessage(string messageId, string flag) => OnFlag(messageId, flag);
        }
    }
}
