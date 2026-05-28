using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class OutlookExportSearchResultsToolTests
    {
        private sealed class FakeSurface : MinimalSurface
        {
            public int Count { get; set; }
            public List<MessageSummary> Hits { get; set; } = new List<MessageSummary>();
            public ExportExcelArgs CapturedExcel { get; set; }
            public int? CapturedMaxResults { get; set; }

            public override int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
                => Count;

            public override SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken))
            {
                CapturedMaxResults = args.MaxResults;
                var page = Hits.Take(args.MaxResults).ToList();
                return new SearchResult { Messages = page, TotalMatches = Count, Truncated = Count > page.Count };
            }

            public override FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken))
            {
                CapturedExcel = args;
                return new FileSavedResult
                {
                    Path = @"C:\Users\x\AppData\Local\OutlookAI\Reports\vendor.xlsx",
                    FileUrl = "file:///C:/Users/x/AppData/Local/OutlookAI/Reports/vendor.xlsx",
                    Format = "xlsx",
                    Bytes = 1234,
                    Filename = "vendor.xlsx",
                };
            }
        }

        private static List<MessageSummary> Make(int n)
        {
            var list = new List<MessageSummary>();
            for (var i = 0; i < n; i++)
                list.Add(new MessageSummary
                {
                    Id = "id" + i, Subject = "S" + i, From = "f" + i + "@x.com",
                    To = new[] { "me@x.com" }, ReceivedAt = DateTimeOffset.UtcNow.AddMinutes(i),
                    Snippet = "snip" + i, HasAttachments = false,
                });
            return list;
        }

        [Fact]
        public void Execute_UnderCeiling_ExportsAllAndReportsNotTruncated()
        {
            var surface = new FakeSurface { Count = 12, Hits = Make(12) };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"IT Creations\",\"columns\":[\"subject\",\"from\",\"snippet\"],\"filename_hint\":\"vendor\"}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("file_saved", (string)obj["result_type"]);
            Assert.Equal(12, (int)obj["exported"]);
            Assert.Equal(12, (int)obj["total_matches"]);
            Assert.False((bool)obj["truncated"]);
            Assert.Equal("vendor.xlsx", (string)obj["filename"]);
            Assert.Equal(3, surface.CapturedExcel.Columns.Count);
            Assert.Equal(12, surface.CapturedExcel.Rows.Count);
        }

        [Fact]
        public void Execute_OverCeiling_CapsAndReportsTruncated()
        {
            var surface = new FakeSurface { Count = 5000, Hits = Make(5000) };
            var tool = new OutlookExportSearchResultsTool(maxRowsOverride: 2000);

            var json = tool.ExecuteAsync(
                "{\"from\":\"x\",\"columns\":[\"subject\"]}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal(2000, (int)obj["exported"]);
            Assert.Equal(5000, (int)obj["total_matches"]);
            Assert.True((bool)obj["truncated"]);
            Assert.Equal(2000, surface.CapturedMaxResults);
        }

        [Fact]
        public void Execute_ZeroMatches_ReturnsNoMatchesWithoutWritingFile()
        {
            var surface = new FakeSurface { Count = 0, Hits = Make(0) };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"nobody\",\"columns\":[\"subject\"]}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("no_matches", (string)obj["result_type"]);
            Assert.Null(surface.CapturedExcel);
        }
    }
}
