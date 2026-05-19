using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.Chat;
using OutlookAI.Tests.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.TaskPane.Chat
{
    public class ExportBridgeTests
    {
        [Fact]
        public async Task ExportPdf_PostsBackFileCardScript()
        {
            var scripts = new List<string>();
            var surface = new Surface { Result = SavedResult() };
            var bridge = new ExportBridge(surface, new Policy(), s => { scripts.Add(s); return Task.CompletedTask; });

            var handled = await bridge.HandleAsync("export_pdf", new JObject(
                new JProperty("message_id", "msg_1"),
                new JProperty("filename_hint", "Report"),
                new JProperty("content_markdown", "# Hello")), CancellationToken.None);

            Assert.True(handled);
            Assert.Equal(1, surface.CallCount);
            Assert.Equal("Report", surface.ObservedArgs.FilenameHint);
            Assert.Equal("# Hello", surface.ObservedArgs.ContentMarkdown);
            Assert.Single(scripts);
            Assert.Contains("outlookai.onFileSaved", scripts[0]);
            Assert.Contains("msg_1", scripts[0]);
            Assert.Contains("Quotes.pdf", scripts[0]);
        }

        [Fact]
        public async Task ExportPdf_OnExceptionPostsErrorCard()
        {
            var scripts = new List<string>();
            var surface = new Surface { Exception = new ExportException("pdf_render_failed", "boom") };
            var bridge = new ExportBridge(surface, new Policy(), s => { scripts.Add(s); return Task.CompletedTask; });

            var handled = await bridge.HandleAsync("export_pdf", new JObject(
                new JProperty("message_id", "msg_2"),
                new JProperty("content_markdown", "body")), CancellationToken.None);

            Assert.True(handled);
            Assert.Single(scripts);
            Assert.Contains("outlookai.onExportError", scripts[0]);
            Assert.Contains("msg_2", scripts[0]);
            Assert.Contains("pdf_render_failed", scripts[0]);
            Assert.Contains("boom", scripts[0]);
        }

        [Fact]
        public async Task OpenFile_ValidatesPathThenLaunches()
        {
            var policy = new Policy();
            var launched = new List<string>();
            var bridge = new ExportBridge(new Surface(), policy, _ => Task.CompletedTask, launched.Add);

            var handled = await bridge.HandleAsync("open_file", new JObject(
                new JProperty("path", @"C:\Reports\Quotes.pdf")), CancellationToken.None);

            Assert.True(handled);
            Assert.Equal(new[] { @"C:\Reports\Quotes.pdf" }, policy.Paths);
            Assert.Equal(new[] { @"C:\Reports\Quotes.pdf" }, launched);
        }

        [Fact]
        public async Task OpenFile_RejectedPathDoesNotLaunchAndPostsError()
        {
            var scripts = new List<string>();
            var launched = new List<string>();
            var policy = new Policy { Exception = new UnauthorizedExportPathException("blocked") };
            var bridge = new ExportBridge(new Surface(), policy, s => { scripts.Add(s); return Task.CompletedTask; }, launched.Add);

            var handled = await bridge.HandleAsync("open_file", new JObject(
                new JProperty("path", @"C:\Windows\cmd.exe")), CancellationToken.None);

            Assert.True(handled);
            Assert.Empty(launched);
            Assert.Single(scripts);
            Assert.Contains("outlookai.onExportError", scripts[0]);
            Assert.Contains("open_file_failed", scripts[0]);
        }

        [Fact]
        public async Task Reveal_ValidatesPathThenLaunches()
        {
            var policy = new Policy();
            var launched = new List<string>();
            var bridge = new ExportBridge(new Surface(), policy, _ => Task.CompletedTask, null, launched.Add);

            var handled = await bridge.HandleAsync("reveal_in_explorer", new JObject(
                new JProperty("path", @"C:\Reports\Quotes.pdf")), CancellationToken.None);

            Assert.True(handled);
            Assert.Equal(new[] { @"C:\Reports\Quotes.pdf" }, policy.Paths);
            Assert.Equal(new[] { @"C:\Reports\Quotes.pdf" }, launched);
        }

        [Fact]
        public async Task UnknownType_NotHandled()
        {
            var bridge = new ExportBridge(new Surface(), new Policy(), _ => Task.CompletedTask);

            Assert.False(await bridge.HandleAsync("nope", new JObject(), CancellationToken.None));
            Assert.False(await bridge.HandleAsync(null, new JObject(), CancellationToken.None));
        }

        [Fact]
        public void Constructor_NullGuards()
        {
            Assert.Throws<ArgumentNullException>(() => new ExportBridge(null, new Policy(), _ => Task.CompletedTask));
            Assert.Throws<ArgumentNullException>(() => new ExportBridge(new Surface(), null, _ => Task.CompletedTask));
            Assert.Throws<ArgumentNullException>(() => new ExportBridge(new Surface(), new Policy(), null));
        }

        private static FileSavedResult SavedResult()
        {
            return new FileSavedResult
            {
                Path = @"C:\Exports\Quotes.pdf",
                FileUrl = "file:///C:/Exports/Quotes.pdf",
                Format = "pdf",
                Bytes = 12345,
                Filename = "Quotes.pdf",
            };
        }

        private sealed class Surface : MinimalSurface
        {
            public FileSavedResult Result { get; set; }
            public Exception Exception { get; set; }
            public ExportPdfArgs ObservedArgs { get; private set; }
            public int CallCount { get; private set; }

            public override FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
            {
                CallCount++;
                ObservedArgs = args;
                if (Exception != null) throw Exception;
                return Result;
            }
        }

        private sealed class Policy : IExportPathPolicy
        {
            private readonly List<string> _paths = new List<string>();

            public Exception Exception { get; set; }
            public IReadOnlyList<string> Paths => _paths;

            public void RequireInsideReportsDir(string path)
            {
                _paths.Add(path);
                if (Exception != null) throw Exception;
            }
        }
    }
}
