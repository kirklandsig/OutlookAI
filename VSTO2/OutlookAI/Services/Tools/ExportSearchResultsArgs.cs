using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Args for outlook_export_search_results: a search filter (same shape as
    /// outlook_search_messages) plus the mechanical columns to project and an
    /// optional filename hint. No body reads — columns come from the search
    /// projection only.
    /// </summary>
    public sealed class ExportSearchResultsArgs
    {
        public SearchMessagesArgs Filter { get; set; }
        public IReadOnlyList<string> Columns { get; set; }
        public string FilenameHint { get; set; }
        public string SheetName { get; set; }
    }
}
