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
        public void Execute_CeilingAboveInteractiveCap_ClampedToInteractiveCap()
        {
            // #12.1: the bulk tool must not produce a workbook larger than the
            // interactive outlook_export_excel tool permits. An admin who sets
            // MaxBulkExportRows between 10,001 and 50,000 would otherwise get a
            // bulk export exceeding ExportExcelArgsParser's 10,000-row cap. The
            // bulk ceiling shares that one cap constant, so a 20,000 override is
            // clamped down to 10,000.
            var surface = new FakeSurface { Count = 20000, Hits = Make(20000) };
            var tool = new OutlookExportSearchResultsTool(maxRowsOverride: 20000);

            var json = tool.ExecuteAsync(
                "{\"from\":\"x\",\"columns\":[\"subject\"]}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal(BulkExportRowCap.Max, surface.CapturedMaxResults);
            Assert.Equal(BulkExportRowCap.Max, (int)obj["exported"]);
            Assert.Equal(20000, (int)obj["total_matches"]);
            Assert.True((bool)obj["truncated"]);
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

        [Fact]
        public void Execute_SanitizesInvalidSheetNameChars_StillExports()
        {
            var surface = new FakeSurface { Count = 3, Hits = Make(3) };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"x\",\"columns\":[\"subject\"],\"filename_hint\":\"Acme: invoices/2025\"}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("file_saved", (string)obj["result_type"]);
            // Sheet name must contain none of the invalid Excel chars and be <= 31 chars.
            var sheet = surface.CapturedExcel.SheetName;
            Assert.DoesNotContain(':', sheet);
            Assert.DoesNotContain('/', sheet);
            Assert.True(sheet.Length <= 31);
        }

        [Fact]
        public void Execute_FilenameHintHistory_AvoidsReservedSheetName()
        {
            // "History" is reserved by Excel; ClosedXML throws if a worksheet is
            // named it. filename_hint defaults to the sheet name, so this must
            // be remapped rather than crash the export.
            var surface = new FakeSurface { Count = 2, Hits = Make(2) };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"x\",\"columns\":[\"subject\"],\"filename_hint\":\"History\"}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("file_saved", (string)obj["result_type"]);
            Assert.NotEqual("History", surface.CapturedExcel.SheetName);
        }

        [Fact]
        public void Execute_ProjectsCellValues_ReceivedAtToJoinAndBool()
        {
            var when = new DateTimeOffset(2026, 3, 4, 5, 6, 7, TimeSpan.Zero);
            var surface = new FakeSurface
            {
                Count = 1,
                Hits = new List<MessageSummary>
                {
                    new MessageSummary
                    {
                        Id = "id1", Subject = "Hello", From = "alice@x.com",
                        To = new[] { "bob@x.com", "carol@x.com" },
                        ReceivedAt = when, Snippet = "hi there", HasAttachments = true,
                    },
                },
            };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"x\",\"columns\":[\"received_at\",\"to\",\"has_attachments\",\"subject\"]}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            JObject.Parse(json);   // ensure valid JSON result

            var rows = surface.CapturedExcel.Rows;
            Assert.Single(rows);
            var cells = rows[0];
            // Column order: received_at, to, has_attachments, subject
            Assert.Equal(when.ToString("o"), (string)cells[0]);
            Assert.Equal("bob@x.com; carol@x.com", (string)cells[1]);
            Assert.Equal(true, (bool)cells[2]);
            Assert.Equal("Hello", (string)cells[3]);
        }

        [Fact]
        public void Execute_FolderColumn_ProjectsFolderName()
        {
            // #12.3: the `folder` column was deferred in v2.1.2 because
            // MessageSummary carried no folder name. Now that it does, a
            // requested `folder` column must render the message's folder.
            var surface = new FakeSurface
            {
                Count = 1,
                Hits = new List<MessageSummary>
                {
                    new MessageSummary
                    {
                        Id = "id1", Subject = "Hello", From = "alice@x.com",
                        To = new[] { "me@x.com" }, ReceivedAt = DateTimeOffset.UtcNow,
                        Snippet = "hi", HasAttachments = false, FolderName = "Clients/Acme",
                    },
                },
            };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync(
                "{\"from\":\"x\",\"columns\":[\"subject\",\"folder\"]}",
                surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("file_saved", (string)obj["result_type"]);
            Assert.Equal(2, surface.CapturedExcel.Columns.Count);
            Assert.Equal("Folder", surface.CapturedExcel.Columns[1].Name);
            // Column order: subject, folder
            Assert.Equal("Clients/Acme", (string)surface.CapturedExcel.Rows[0][1]);
        }

        [Fact]
        public void Execute_DefaultColumns_WhenNoneProvided()
        {
            var surface = new FakeSurface { Count = 2, Hits = Make(2) };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync("{\"from\":\"x\"}", surface, CancellationToken.None).GetAwaiter().GetResult();
            JObject.Parse(json);

            // DefaultColumns = received_at, from, to, subject, snippet => 5 columns
            Assert.Equal(5, surface.CapturedExcel.Columns.Count);
        }

        [Fact]
        public void Execute_MalformedJson_ReturnsInvalidArgs()
        {
            var surface = new FakeSurface { Count = 0 };
            var tool = new OutlookExportSearchResultsTool();

            var json = tool.ExecuteAsync("{not valid json", surface, CancellationToken.None).GetAwaiter().GetResult();
            var obj = JObject.Parse(json);

            Assert.Equal("invalid_args", (string)obj["error"]["code"]);
        }
    }
}
