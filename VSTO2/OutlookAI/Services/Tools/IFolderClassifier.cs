using System;
using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Decides whether a folder is an Outlook system / noise folder that
    /// must never produce results for mailbox search. Centralised so both
    /// the AdvancedSearch projection and the iterative fallback honour the
    /// same skip rules.
    /// </summary>
    public interface IFolderClassifier
    {
        bool IsSystemFolder(string folderName, bool defaultItemTypeIsMail);
    }

    public sealed class FolderClassifier : IFolderClassifier
    {
        // Real Outlook folder display names. "Junk E-mail" uses a hyphen in
        // modern Outlook; "Junk Email" appears in older locales. Both skip.
        private static readonly HashSet<string> _systemNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Deleted Items",
                "Junk E-mail",
                "Junk Email",
                "Drafts",
                "Outbox",
                "Sync Issues",
                "Sync Issues (This computer only)",
                "Conflicts",
                "Local Failures",
                "Server Failures",
                "RSS Feeds",
                "RSS Subscriptions",
                "Conversation Action Settings",
                "Conversation History",
                "Quick Step Settings",
                "News Feed",
                "Feeds",
                "Files",
                "Detected Items",
                "Working Set",
                "Yammer Root",
            };

        public bool IsSystemFolder(string folderName, bool defaultItemTypeIsMail)
        {
            if (!defaultItemTypeIsMail) return true;
            if (string.IsNullOrWhiteSpace(folderName)) return true;
            return _systemNames.Contains(folderName.Trim());
        }
    }
}
