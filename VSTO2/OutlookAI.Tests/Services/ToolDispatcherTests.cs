using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class ToolDispatcherTests
    {
        private sealed class StubTool : IOutlookTool
        {
            public string Name => "outlook_stub";
            public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
                => Task.FromResult("{\"ok\":true,\"echo\":" + argsJson + "}");
        }

        private sealed class ThrowingTool : IOutlookTool
        {
            public string Name => "outlook_boom";
            public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
                => throw new System.InvalidOperationException("kaboom");
        }

        [Fact]
        public async Task DispatchAsync_RoutesByName()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new StubTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_stub", "{\"x\":1}", CancellationToken.None);
            Assert.Contains("\"ok\":true", result);
            Assert.Contains("\"x\":1", result);
        }

        [Fact]
        public async Task DispatchAsync_ReturnsStructuredErrorForUnknownTool()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new StubTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_missing", "{}", CancellationToken.None);
            Assert.Contains("\"code\":\"unknown_tool\"", result);
        }

        [Fact]
        public async Task DispatchAsync_ReturnsStructuredErrorForMalformedJson()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new StubTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_stub", "{ not json", CancellationToken.None);
            Assert.Contains("\"code\":\"invalid_arguments\"", result);
        }

        [Fact]
        public async Task DispatchAsync_WrapsToolExceptionAsStructuredError()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new ThrowingTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_boom", "{}", CancellationToken.None);
            Assert.Contains("\"code\":\"InvalidOperationException\"", result);
            Assert.Contains("kaboom", result);
        }

        [Fact]
        public async Task DispatchAsync_AcceptsEmptyArgsJson()
        {
            var d = new ToolDispatcher(new IOutlookTool[] { new StubTool() }, surface: null);
            var result = await d.DispatchAsync("outlook_stub", "", CancellationToken.None);
            Assert.Contains("\"ok\":true", result);
        }
    }
}
