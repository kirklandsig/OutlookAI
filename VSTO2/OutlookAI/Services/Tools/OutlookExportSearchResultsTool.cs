using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// outlook_export_search_results. Deterministic, complete (up to a
    /// configurable ceiling) mechanical Excel export of a search. Counts the
    /// true total, collects up to the ceiling via the fast Table-API search
    /// path (no per-message body reads), projects the requested mechanical
    /// columns, and reports "exported N of M". For AI-synthesized exports
    /// ("summarize each email") the model must still read + accumulate; this
    /// tool only emits raw projected fields.
    /// </summary>
    public sealed class OutlookExportSearchResultsTool : IOutlookTool
    {
        private readonly int? _maxRowsOverride;

        public OutlookExportSearchResultsTool() : this(null) { }

        // Test seam so the ceiling can be set without touching global Config.
        internal OutlookExportSearchResultsTool(int? maxRowsOverride)
        {
            _maxRowsOverride = maxRowsOverride;
        }

        public string Name => "outlook_export_search_results";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var args = ExportSearchResultsArgsParser.Parse(argsJson);
                var ceiling = _maxRowsOverride ?? Config.MaxBulkExportRows;
                if (ceiling < 1) ceiling = 1;
                // Never exceed the cap the interactive outlook_export_excel
                // path enforces, so the two exporters can't produce workbooks
                // of differing maximum size (#12.1).
                if (ceiling > BulkExportRowCap.Max) ceiling = BulkExportRowCap.Max;

                // Two separate Outlook traversals: count first (true total M),
                // then collect (up to the ceiling, N). On a churning mailbox a
                // message can arrive/move/delete between the two calls, and the
                // count path's per-folder cap (SearchFallbackBudget.CountModeCap,
                // 5000) can differ from the search path's per-folder budget — so
                // the reported "exported N of M" / truncated can be slightly off.
                // Accepted by design (#12.2): a single fused traversal would lose
                // the accurate M denominator (you'd only know M up to the
                // ceiling). N is always <= M here because we set MaxResults to
                // min(ceiling, total) and SearchResultProjector clamps its output
                // globally to MaxResults (not per-folder), so the report can never
                // claim it exported more than the total.
                // True total (count-mode, bounded per folder).
                var total = surface.CountMessages(args.Filter, ct);
                if (total <= 0)
                {
                    return Task.FromResult(new JObject(
                        new JProperty("result_type", "no_matches"),
                        new JProperty("total_matches", 0),
                        new JProperty("message", "No messages matched the filter; nothing was exported."))
                        .ToString(Newtonsoft.Json.Formatting.None));
                }

                // Collect up to the ceiling via the normal search path.
                args.Filter.MaxResults = Math.Min(ceiling, total);
                var result = surface.SearchMessages(args.Filter, ct) ?? new SearchResult();
                var hits = result.Messages ?? new MessageSummary[0];

                var excelArgs = BuildExcelArgs(args, hits);
                var saved = surface.ExportExcel(excelArgs, ct);

                var exported = hits.Count;
                var truncated = exported < total;

                return Task.FromResult(new JObject(
                    new JProperty("result_type", "file_saved"),
                    new JProperty("path", saved.Path ?? ""),
                    new JProperty("file_url", saved.FileUrl ?? ""),
                    new JProperty("format", saved.Format ?? ""),
                    new JProperty("bytes", saved.Bytes),
                    new JProperty("filename", saved.Filename ?? ""),
                    new JProperty("exported", exported),
                    new JProperty("total_matches", total),
                    new JProperty("truncated", truncated))
                    .ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Export cancelled by user."));
            }
            catch (ToolArgValidationException ex)
            {
                return Task.FromResult(BuildError(ex.Code, ex.Message));
            }
            catch (ExportException ex)
            {
                return Task.FromResult(BuildError(ex.Code, ex.Message));
            }
            catch (Exception ex)
            {
                return Task.FromResult(BuildError("export_failed", ex.Message));
            }
        }

        private static ExportExcelArgs BuildExcelArgs(ExportSearchResultsArgs args, IReadOnlyList<MessageSummary> hits)
        {
            var columns = new List<ExcelColumnSpec>();
            foreach (var c in args.Columns)
            {
                columns.Add(new ExcelColumnSpec { Name = HeaderFor(c), Type = TypeFor(c) });
            }

            var rows = new List<JToken[]>(hits.Count);
            foreach (var m in hits)
            {
                var cells = new JToken[args.Columns.Count];
                for (var i = 0; i < args.Columns.Count; i++)
                {
                    cells[i] = CellFor(args.Columns[i], m);
                }
                rows.Add(cells);
            }

            return new ExportExcelArgs
            {
                FilenameHint = args.FilenameHint,
                SheetName = SanitizeSheetName(args.SheetName),
                Columns = columns,
                Rows = rows,
            };
        }

        private static string HeaderFor(string col)
        {
            switch (col)
            {
                case "subject": return "Subject";
                case "from": return "From";
                case "to": return "To";
                case "received_at": return "Received";
                case "snippet": return "Snippet";
                case "has_attachments": return "Has Attachments";
                case "folder": return "Folder";
                default: return col;
            }
        }

        private static ExcelColumnType TypeFor(string col)
        {
            switch (col)
            {
                case "received_at": return ExcelColumnType.DateTime;
                case "has_attachments": return ExcelColumnType.Boolean;
                default: return ExcelColumnType.Text;
            }
        }

        private static JToken CellFor(string col, MessageSummary m)
        {
            switch (col)
            {
                case "subject": return m.Subject ?? "";
                case "from": return m.From ?? "";
                case "to": return string.Join("; ", m.To ?? new string[0]);
                case "received_at": return m.ReceivedAt.ToString("o");
                case "snippet": return m.Snippet ?? "";
                case "has_attachments": return m.HasAttachments;
                case "folder": return m.FolderName ?? "";
                default: return "";
            }
        }

        private static readonly char[] InvalidSheetNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

        private static string SanitizeSheetName(string raw)
        {
            var s = string.IsNullOrWhiteSpace(raw) ? "Results" : raw;
            foreach (var c in InvalidSheetNameChars) s = s.Replace(c, ' ');
            s = s.Trim().Trim('\'').Trim();
            if (string.IsNullOrEmpty(s)) s = "Results";
            // "History" is reserved by Excel; ClosedXML throws if a worksheet
            // is named it (case-insensitive). filename_hint defaults to the
            // sheet name, so a model-supplied "History" would otherwise crash
            // the export with a misleading error.
            if (string.Equals(s, "History", StringComparison.OrdinalIgnoreCase)) s = "Results";
            return s.Length <= 31 ? s : s.Substring(0, 31);
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
