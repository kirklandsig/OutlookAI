using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookExportPdfToolTests
    {
        [Fact]
        public void Name_IsOutlookExportPdf()
        {
            var tool = new OutlookExportPdfTool();

            Assert.Equal("outlook_export_pdf", tool.Name);
        }

        [Fact]
        public async Task Execute_ReturnsFileSavedEnvelopeOnSuccess()
        {
            var surface = new Surface
            {
                Result = new FileSavedResult
                {
                    Path = @"C:\Exports\Quotes.pdf",
                    FileUrl = "file:///C:/Exports/Quotes.pdf",
                    Format = "pdf",
                    Bytes = 12345,
                    Filename = "Quotes.pdf",
                }
            };
            var tool = new OutlookExportPdfTool();

            var json = await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            var result = JObject.Parse(json);
            Assert.Equal("file_saved", (string)result["result_type"]);
            Assert.Equal(@"C:\Exports\Quotes.pdf", (string)result["path"]);
            Assert.Equal("file:///C:/Exports/Quotes.pdf", (string)result["file_url"]);
            Assert.Equal("pdf", (string)result["format"]);
            Assert.Equal(12345L, (long)result["bytes"]);
            Assert.Equal("Quotes.pdf", (string)result["filename"]);
        }

        [Fact]
        public async Task Execute_CapturesParsedArgs()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportPdfTool();

            await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            Assert.Equal(1, surface.CallCount);
            Assert.NotNull(surface.ObservedArgs);
            Assert.Equal("Quotes", surface.ObservedArgs.FilenameHint);
            Assert.Equal("# Budget", surface.ObservedArgs.ContentMarkdown);
            Assert.Equal("Budget Report", surface.ObservedArgs.Title);
            Assert.Equal("Q1", surface.ObservedArgs.Subtitle);
        }

        [Fact]
        public async Task Execute_InvalidArgs_ReturnsValidationErrorEnvelope()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportPdfTool();

            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("invalid_args", (string)error["code"]);
            Assert.Contains("content_markdown is required", (string)error["message"]);
            Assert.Equal(0, surface.CallCount);
        }

        [Fact]
        public async Task Execute_ContentTooLarge_ReturnsContentTooLargeEnvelope()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportPdfTool();
            var content = new string('x', 250001);
            var args = new JObject(new JProperty("content_markdown", content));

            var json = await tool.ExecuteAsync(args.ToString(Newtonsoft.Json.Formatting.None), surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("content_too_large", (string)error["code"]);
            Assert.Contains("content_markdown must be <= 250000 characters", (string)error["message"]);
            Assert.Equal(0, surface.CallCount);
        }

        [Fact]
        public async Task Execute_ExportException_ReturnsExportErrorEnvelope()
        {
            var surface = new Surface { Exception = new ExportException("webview2_missing", "install runtime") };
            var tool = new OutlookExportPdfTool();

            var json = await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("webview2_missing", (string)error["code"]);
            Assert.Equal("install runtime", (string)error["message"]);
        }

        [Fact]
        public async Task Execute_PreCancelledToken_ReturnsCancelledEnvelopeAndDoesNotCallSurface()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportPdfTool();
            using (var cts = new CancellationTokenSource())
            {
                cts.Cancel();

                var json = await tool.ExecuteAsync(ValidArgsJson(), surface, cts.Token);

                var error = JObject.Parse(json)["error"];
                Assert.Equal("cancelled", (string)error["code"]);
                Assert.Equal(0, surface.CallCount);
            }
        }

        [Fact]
        public async Task Execute_UnexpectedSurfaceException_ReturnsGenericExportFailureEnvelope()
        {
            var surface = new Surface { Exception = new InvalidOperationException("boom") };
            var tool = new OutlookExportPdfTool();

            var json = await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("pdf_render_failed", (string)error["code"]);
            Assert.Equal("boom", (string)error["message"]);
        }

        private static string ValidArgsJson()
        {
            return "{"
                + "\"filename_hint\":\"Quotes\","
                + "\"content_markdown\":\"# Budget\","
                + "\"title\":\"Budget Report\","
                + "\"subtitle\":\"Q1\"}";
        }

        private static FileSavedResult SavedResult()
        {
            return new FileSavedResult
            {
                Path = @"C:\Exports\Quotes.pdf",
                FileUrl = "file:///C:/Exports/Quotes.pdf",
                Format = "pdf",
                Bytes = 1,
                Filename = "Quotes.pdf",
            };
        }

        private sealed class Surface : MinimalSurface
        {
            public FileSavedResult Result { get; set; }
            public Exception Exception { get; set; }
            public ExportPdfArgs ObservedArgs { get; private set; }
            public CancellationToken ObservedCt { get; private set; }
            public int CallCount { get; private set; }

            public override FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken))
            {
                ct.ThrowIfCancellationRequested();
                CallCount++;
                ObservedArgs = args;
                ObservedCt = ct;
                if (Exception != null) throw Exception;
                return Result;
            }
        }
    }
}
