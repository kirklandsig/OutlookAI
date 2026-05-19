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
        public const int MaxListFolders = 200;
        public const int MaxSearchFolders = 5000;

        // Tight floor. For oldest/newest queries with MaxResults=1, taking
        // exactly ONE item per folder is correct: the global sort across
        // all folders picks the absolute oldest/newest. Each additional
        // per-folder item costs ~100-300 ms of Outlook COM property
        // access on uncached items, so over-provisioning here adds tens
        // of seconds of pointless work on big mailboxes.
        //
        // The CollectFolderInputs loop already filters non-MailItems
        // (continue without incrementing taken), so taking N == MaxResults
        // MailItems is exact.
        //
        // The classifier is applied at the folder level (whole system
        // folders are skipped before CollectFolderInputs is called), so
        // no per-item classifier buffer is needed here.
        private const int Floor = 1;

        // Hard cap used in count mode (MaxResults = int.MaxValue) so
        // CountMessages still bounds per-folder work. Counts of bigger
        // folders truncate at this cap — better truncated than frozen.
        private const int CountModeCap = 5000;

        public static int PerFolderItems(SearchMessagesArgs args)
        {
            if (args == null) return Floor;
            if (args.MaxResults == int.MaxValue) return CountModeCap;
            if (args.MaxResults < Floor) return Floor;
            return args.MaxResults;
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

        public static int MaxFoldersForSearch(SearchMessagesArgs args, bool allMail)
        {
            if (!IsBroadFiniteNewestAllMail(args, allMail)) return MaxSearchFolders;
            return MaxListFolders;
        }

        public static SearchMessagesArgs CountFolderResolutionArgs(SearchMessagesArgs args)
        {
            return new SearchMessagesArgs
            {
                FolderId = args?.FolderId,
                MaxResults = int.MaxValue,
            };
        }

        public static bool ShouldStopRecipientAllMailScan(
            SearchMessagesArgs args, string scopeMode, int candidateCount)
        {
            if (args == null) return false;
            if (candidateCount < args.MaxResults) return false;
            if (args.MaxResults <= 0 || args.MaxResults == int.MaxValue) return false;
            if (string.IsNullOrWhiteSpace(args.To)) return false;
            if (!DescendingForSortOrder(args.SortOrder)) return false;
            if (!string.IsNullOrWhiteSpace(args.FolderId)) return false;

            return string.Equals(scopeMode, "all_mail", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scopeMode, "auto", StringComparison.OrdinalIgnoreCase);
        }

        public static bool ShouldStopBroadAllMailScan(
            SearchMessagesArgs args, string scopeMode, int candidateCount)
        {
            if (args == null) return false;
            if (candidateCount < args.MaxResults) return false;
            return IsBroadFiniteNewestAllMail(args, IsAllMailScope(scopeMode));
        }

        private static bool IsBroadFiniteNewestAllMail(SearchMessagesArgs args, bool allMail)
        {
            if (args == null) return false;
            if (!allMail) return false;
            if (args.MaxResults <= 0 || args.MaxResults == int.MaxValue) return false;
            if (!DescendingForSortOrder(args.SortOrder)) return false;
            if (!string.IsNullOrWhiteSpace(args.FolderId)) return false;
            if (!IsAnyFilter(args.AttachmentFilter)) return false;
            if (!IsAnyFilter(args.ReadStatus)) return false;
            if (!IsAnyFilter(args.FlagStatus)) return false;
            if (!IsAnyFilter(args.ImportanceFilter)) return false;
            if (args.HasAttachment == true) return false;
            if (args.IsUnread == true) return false;
            if (args.IsFlagged == true) return false;
            if (!string.IsNullOrWhiteSpace(args.Importance)) return false;

            return string.IsNullOrWhiteSpace(args.Query)
                && string.IsNullOrWhiteSpace(args.From)
                && string.IsNullOrWhiteSpace(args.To)
                && string.IsNullOrWhiteSpace(args.SubjectContains)
                && string.IsNullOrWhiteSpace(args.BodyContains);
        }

        private static bool IsAnyFilter(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                || string.Equals(value, "any", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsAllMailScope(string scopeMode)
        {
            return string.Equals(scopeMode, "all_mail", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scopeMode, "auto", StringComparison.OrdinalIgnoreCase);
        }
    }
}
