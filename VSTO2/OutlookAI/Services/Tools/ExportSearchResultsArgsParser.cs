using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Parses outlook_export_search_results args. Reuses SearchMessagesArgsParser
    /// for the filter (count-mode shape — no max_results clamp), validates the
    /// requested columns against a fixed allow-list of mechanical fields, drops
    /// unknown / duplicate columns, and falls back to DefaultColumns when none
    /// survive.
    /// </summary>
    internal static class ExportSearchResultsArgsParser
    {
        public static readonly string[] AllowedColumns =
        {
            "subject", "from", "to", "received_at", "snippet", "has_attachments", "folder",
        };

        public static readonly string[] DefaultColumns =
        {
            "received_at", "from", "to", "subject", "snippet",
        };

        private const string DefaultFilenameHint = "OutlookAI-Search-Export";

        public static ExportSearchResultsArgs Parse(string argsJson)
        {
            var raw = string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson;

            // Reuse the existing filter parser (count shape: no max_results cap).
            var filter = SearchMessagesArgsParser.ParseCount(raw);

            var obj = JObject.Parse(raw);

            var columns = ParseColumns(obj["columns"]);
            var filenameHint = CleanString(obj["filename_hint"]) ?? DefaultFilenameHint;
            var sheetName = CleanString(obj["sheet_name"]) ?? filenameHint;

            return new ExportSearchResultsArgs
            {
                Filter = filter,
                Columns = columns,
                FilenameHint = filenameHint,
                SheetName = sheetName,
            };
        }

        private static IReadOnlyList<string> ParseColumns(JToken token)
        {
            var arr = token as JArray;
            if (arr == null) return DefaultColumns;

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ordered = new List<string>();
            foreach (var t in arr)
            {
                var name = (t?.Type == JTokenType.String) ? ((string)t)?.Trim().ToLowerInvariant() : null;
                if (string.IsNullOrEmpty(name)) continue;
                if (!AllowedColumns.Contains(name)) continue;
                if (!seen.Add(name)) continue;
                ordered.Add(name);
            }

            return ordered.Count > 0 ? (IReadOnlyList<string>)ordered : DefaultColumns;
        }

        private static string CleanString(JToken token)
        {
            if (token == null || token.Type != JTokenType.String) return null;
            var v = ((string)token)?.Trim();
            return string.IsNullOrEmpty(v) ? null : v;
        }
    }
}
