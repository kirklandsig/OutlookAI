using System.Collections.Generic;
using System.Text;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Formats a list of MAPI folder paths into the comma-separated,
    /// single-quoted scope string that <c>Application.AdvancedSearch</c>
    /// expects, escaping embedded single quotes per Outlook DASL rules.
    /// </summary>
    public static class SearchScopeFormatter
    {
        public static string Format(IEnumerable<string> folderPaths)
        {
            if (folderPaths == null) return "";
            var sb = new StringBuilder();
            bool first = true;
            foreach (var raw in folderPaths)
            {
                if (string.IsNullOrWhiteSpace(raw)) continue;
                if (!first) sb.Append(',');
                first = false;
                sb.Append('\'').Append(raw.Replace("'", "''")).Append('\'');
            }
            return sb.ToString();
        }
    }

    /// <summary>
    /// Result of <see cref="SearchScopeFormatter"/>-aware scope resolution
    /// in LiveOutlookSurface. The scope string is what AdvancedSearch
    /// consumes; SearchSubFolders controls whether each scope entry is
    /// expanded recursively; ResolvedFolderPaths is kept for trace logging.
    /// </summary>
    public sealed class SearchScope
    {
        public string ScopeString { get; set; }
        public bool SearchSubFolders { get; set; }
        public IReadOnlyList<string> ResolvedFolderPaths { get; set; }
    }
}
