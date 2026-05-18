using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Pure projection helper. Filters out system folders via
    /// IFolderClassifier, sorts by ReceivedAt per args.SortOrder, clamps to
    /// args.MaxResults, then evaluates each survivor's SnippetFactory.
    /// Resilient: a throwing SnippetFactory yields an empty snippet but
    /// does not poison the batch.
    /// </summary>
    public static class SearchResultProjector
    {
        public static IReadOnlyList<MessageSummary> Project(
            IEnumerable<MessageProjectionInput> items,
            SearchMessagesArgs args,
            IFolderClassifier classifier)
        {
            items = items ?? new MessageProjectionInput[0];
            args = args ?? new SearchMessagesArgs();
            classifier = classifier ?? new FolderClassifier();

            var filtered = items.Where(i =>
                i != null
                && !classifier.IsSystemFolder(i.FolderName, i.FolderDefaultItemTypeIsMail));

            var ordered = string.Equals(args.SortOrder, "oldest", StringComparison.OrdinalIgnoreCase)
                ? filtered.OrderBy(i => i.ReceivedAt)
                : filtered.OrderByDescending(i => i.ReceivedAt);

            var maxResults = args.MaxResults > 0 ? args.MaxResults : 25;
            var top = ordered.Take(maxResults).ToList();

            var output = new List<MessageSummary>(top.Count);
            foreach (var i in top)
            {
                string snippet = "";
                try { snippet = i.SnippetFactory != null ? (i.SnippetFactory() ?? "") : ""; }
                catch { snippet = ""; }

                output.Add(new MessageSummary
                {
                    Id = i.Id ?? "",
                    Subject = i.Subject ?? "",
                    From = i.From ?? "",
                    To = i.To ?? new string[0],
                    ReceivedAt = i.ReceivedAt,
                    Snippet = snippet,
                    HasAttachments = i.HasAttachments,
                });
            }
            return output;
        }
    }
}
