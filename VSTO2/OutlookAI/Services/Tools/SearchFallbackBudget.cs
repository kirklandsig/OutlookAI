using System;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Phase 3b hotfix. Bounds the per-folder iteration in
    /// LiveOutlookSurface's iterative fallback so a single huge folder
    /// can't lock the Outlook UI thread for tens of seconds.
    ///
    /// Trace data that motivated this: a 588-item folder took ~63s with
    /// ~107 ms per Outlook MailItem property access. Without a per-folder
    /// cap, enumerating every item in every folder froze Outlook for 2+
    /// minutes. With a cap of MaxResults * multiplier, each folder costs
    /// at most a few hundred ms.
    /// </summary>
    public static class SearchFallbackBudget
    {
        // Floor: even for MaxResults=1 we take a few extra to leave room
        // for the projector's classifier filter to drop a few system-folder
        // items before clamping to the user's requested top-N globally.
        private const int Floor = 5;

        // Multiplier: take this many candidate items per folder relative
        // to the user's MaxResults. Five is enough that the global
        // classifier filter + sort + clamp can produce a clean top-N from
        // a buffered pool without being burnt by edge cases (system
        // subfolders inside an otherwise mail-typed folder, etc.).
        private const int Multiplier = 5;

        // Hard cap used in count mode (MaxResults = int.MaxValue) so
        // CountMessages still bounds per-folder work. Counts of bigger
        // folders truncate at this cap — better truncated than frozen.
        private const int CountModeCap = 5000;

        public static int PerFolderItems(SearchMessagesArgs args)
        {
            if (args == null) return Floor;
            if (args.MaxResults == int.MaxValue) return CountModeCap;
            if (args.MaxResults < 1) return Floor;
            return Math.Max(Floor, args.MaxResults * Multiplier);
        }

        /// <summary>
        /// True when the per-folder Items.Sort call should be descending
        /// (newest first); false for "oldest" sort. Mirrors
        /// SearchResultProjector's interpretation.
        /// </summary>
        public static bool DescendingForSortOrder(string sortOrder)
        {
            return !string.Equals(sortOrder, "oldest", StringComparison.OrdinalIgnoreCase);
        }
    }
}
