using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookExportExcelToolTests
    {
        [Fact]
        public void Name_IsOutlookExportExcel()
        {
            var tool = new OutlookExportExcelTool();

            Assert.Equal("outlook_export_excel", tool.Name);
        }

        [Fact]
        public async Task Execute_ReturnsFileSavedEnvelopeOnSuccess()
        {
            var surface = new Surface
            {
                Result = new FileSavedResult
                {
                    Path = @"C:\Exports\Quotes.xlsx",
                    FileUrl = "file:///C:/Exports/Quotes.xlsx",
                    Format = "xlsx",
                    Bytes = 12345,
                    Filename = "Quotes.xlsx",
                }
            };
            var tool = new OutlookExportExcelTool();

            var json = await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            var result = JObject.Parse(json);
            Assert.Equal("file_saved", (string)result["result_type"]);
            Assert.Equal(@"C:\Exports\Quotes.xlsx", (string)result["path"]);
            Assert.Equal("file:///C:/Exports/Quotes.xlsx", (string)result["file_url"]);
            Assert.Equal("xlsx", (string)result["format"]);
            Assert.Equal(12345L, (long)result["bytes"]);
            Assert.Equal("Quotes.xlsx", (string)result["filename"]);
        }

        [Fact]
        public async Task Execute_CapturesParsedArgsAndPreservesRows()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportExcelTool();

            await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            Assert.Equal(1, surface.CallCount);
            Assert.NotNull(surface.ObservedArgs);
            Assert.Equal("Quotes", surface.ObservedArgs.FilenameHint);
            Assert.Single(surface.ObservedArgs.Columns);
            Assert.Equal("Subject", surface.ObservedArgs.Columns[0].Name);
            Assert.Single(surface.ObservedArgs.Rows);
            Assert.Equal("Budget", (string)surface.ObservedArgs.Rows[0][0]);
        }

        [Fact]
        public async Task Execute_InvalidArgs_ReturnsValidationErrorEnvelope()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportExcelTool();

            var json = await tool.ExecuteAsync("{}", surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("invalid_args", (string)error["code"]);
            Assert.Contains("columns is required", (string)error["message"]);
            Assert.Equal(0, surface.CallCount);
        }

        [Fact]
        public async Task Execute_TooManyRows_ReturnsTooManyRowsEnvelope()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportExcelTool();
            var rows = new JArray();
            for (var i = 0; i < 10001; i++)
            {
                rows.Add(new JArray("row " + i));
            }

            var args = new JObject(
                new JProperty("columns", new JArray(new JObject(
                    new JProperty("name", "Subject"),
                    new JProperty("type", "text")))),
                new JProperty("rows", rows));

            var json = await tool.ExecuteAsync(args.ToString(Newtonsoft.Json.Formatting.None), surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("too_many_rows", (string)error["code"]);
            Assert.Contains("rows count", (string)error["message"]);
            Assert.Equal(0, surface.CallCount);
        }

        [Fact]
        public async Task Execute_MalformedJson_ReturnsInvalidArgsEnvelope()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportExcelTool();

            var json = await tool.ExecuteAsync("{\"columns\":", surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("invalid_args", (string)error["code"]);
            Assert.Contains("Invalid JSON args", (string)error["message"]);
            Assert.Equal(0, surface.CallCount);
        }

        [Fact]
        public async Task Execute_ExportException_ReturnsExportErrorEnvelope()
        {
            var surface = new Surface { Exception = new ExportException("file_locked", "in use by Excel") };
            var tool = new OutlookExportExcelTool();

            var json = await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("file_locked", (string)error["code"]);
            Assert.Equal("in use by Excel", (string)error["message"]);
        }

        [Fact]
        public async Task Execute_PreCancelledToken_ReturnsCancelledEnvelopeAndDoesNotCallSurface()
        {
            var surface = new Surface { Result = SavedResult() };
            var tool = new OutlookExportExcelTool();
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
            var tool = new OutlookExportExcelTool();

            var json = await tool.ExecuteAsync(ValidArgsJson(), surface, CancellationToken.None);

            var error = JObject.Parse(json)["error"];
            Assert.Equal("excel_build_failed", (string)error["code"]);
            Assert.Equal("boom", (string)error["message"]);
        }

        private static string ValidArgsJson()
        {
            return "{"
                + "\"filename_hint\":\"Quotes\","
                + "\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}],"
                + "\"rows\":[[\"Budget\"]]}";
        }

        private static FileSavedResult SavedResult()
        {
            return new FileSavedResult
            {
                Path = @"C:\Exports\Quotes.xlsx",
                FileUrl = "file:///C:/Exports/Quotes.xlsx",
                Format = "xlsx",
                Bytes = 1,
                Filename = "Quotes.xlsx",
            };
        }

        private sealed class Surface : MinimalSurface
        {
            public FileSavedResult Result { get; set; }
            public Exception Exception { get; set; }
            public ExportExcelArgs ObservedArgs { get; private set; }
            public CancellationToken ObservedCt { get; private set; }
            public int CallCount { get; private set; }

            public override FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
            {
                CallCount++;
                ObservedArgs = args;
                ObservedCt = ct;
                if (Exception != null) throw Exception;
                return Result;
            }
        }
    }
}
