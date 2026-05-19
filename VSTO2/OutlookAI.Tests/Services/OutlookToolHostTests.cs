using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using OutlookAI.Tests.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class OutlookToolHostTests
    {
        [Fact]
        public async Task DispatchAsync_RoutesExcelExportTool()
        {
            var surface = new Surface();
            var host = new OutlookToolHost(surface, includeWriteTools: false);

            var json = await host.DispatchAsync("outlook_export_excel", ValidArgsJson(), CancellationToken.None);

            var result = JObject.Parse(json);
            Assert.Equal("file_saved", (string)result["result_type"]);
            Assert.Equal(1, surface.ExportExcelCallCount);
        }

        private static string ValidArgsJson()
        {
            return "{"
                + "\"filename_hint\":\"Quotes\","
                + "\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}],"
                + "\"rows\":[[\"Budget\"]]}";
        }

        private sealed class Surface : MinimalSurface
        {
            public int ExportExcelCallCount { get; private set; }

            public override FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
            {
                ExportExcelCallCount++;
                return new FileSavedResult
                {
                    Path = @"C:\Exports\Quotes.xlsx",
                    FileUrl = "file:///C:/Exports/Quotes.xlsx",
                    Format = "xlsx",
                    Bytes = 1,
                    Filename = "Quotes.xlsx",
                };
            }
        }
    }
}
