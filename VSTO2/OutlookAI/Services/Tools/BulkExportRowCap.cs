namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Single source of truth for the maximum number of rows any Excel export
    /// path may produce. Shared by the interactive <c>outlook_export_excel</c>
    /// tool (<see cref="ExportExcelArgsParser"/>) and the bulk
    /// <c>outlook_export_search_results</c> tool so the two paths can never
    /// drift: the bulk ceiling (<c>Config.MaxBulkExportRows</c>) clamps to this,
    /// and the interactive parser rejects row arrays larger than this.
    /// </summary>
    internal static class BulkExportRowCap
    {
        /// <summary>
        /// Hard upper bound on exported rows across all Excel export paths.
        /// </summary>
        public const int Max = 10000;
    }
}
