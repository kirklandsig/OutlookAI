using System;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    internal static class SearchMessagesArgsParser
    {
        public static SearchMessagesArgs ParseSearch(string argsJson)
        {
            var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);
            var maxResults = args["max_results"]?.Value<int>() ?? 25;
            if (maxResults < 1) maxResults = 1;
            if (maxResults > 100) maxResults = 100;
            return Parse(args, maxResults);
        }

        public static SearchMessagesArgs ParseCount(string argsJson)
        {
            var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);
            return Parse(args, int.MaxValue);
        }

        private static SearchMessagesArgs Parse(JObject args, int maxResults)
        {
            var search = new SearchMessagesArgs
            {
                Query = Clean(args["query"]),
                From = Clean(args["from"]),
                To = Clean(args["to"]) ?? Clean(args["recipient"]),
                SubjectContains = Clean(args["subject_contains"]),
                BodyContains = Clean(args["body_contains"]),
                FolderId = Clean(args["folder_id"]),
                DateFrom = ParseDate(args["date_from"], isUpperBound: false),
                DateTo = ParseDate(args["date_to"], isUpperBound: true),
                MaxResults = maxResults,
                Scope = EnumOrDefault(args["scope"], "auto", "current_folder", "all_mail", "auto"),
                SortOrder = EnumOrDefault(args["sort_order"], "newest", "newest", "oldest"),
                AttachmentFilter = EnumOrDefault(args["attachment_filter"], "any", "any", "with", "without"),
                ReadStatus = EnumOrDefault(args["read_status"], "any", "any", "read", "unread"),
                FlagStatus = EnumOrDefault(args["flag_status"], "any", "any", "flagged", "unflagged"),
                ImportanceFilter = EnumOrDefault(args["importance_filter"], "any", "any", "low", "normal", "high"),
            };

            // Hidden old-shape compatibility: preserve true values only.
            // False was a model default in real traces and must not mean
            // "without/read/unflagged" unless the new tri-state says so.
            if (args["has_attachment"]?.Type == JTokenType.Boolean
                && args["has_attachment"].Value<bool>())
            {
                search.HasAttachment = true;
            }
            if (args["is_unread"]?.Type == JTokenType.Boolean
                && args["is_unread"].Value<bool>())
            {
                search.IsUnread = true;
            }
            if (args["is_flagged"]?.Type == JTokenType.Boolean
                && args["is_flagged"].Value<bool>())
            {
                search.IsFlagged = true;
            }
            var oldImportance = Clean(args["importance"]);
            if (oldImportance != null)
            {
                oldImportance = oldImportance.ToLowerInvariant();
                if (oldImportance == "low" || oldImportance == "high")
                {
                    search.Importance = oldImportance;
                }
            }

            return search;
        }

        private static string Clean(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            var value = ((string)token)?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static string EnumOrDefault(JToken token, string fallback, params string[] allowed)
        {
            var value = Clean(token);
            if (value == null) return fallback;
            value = value.ToLowerInvariant();
            foreach (var allowedValue in allowed)
            {
                if (value == allowedValue) return value;
            }
            return fallback;
        }

        private static DateTimeOffset? ParseDate(JToken token, bool isUpperBound)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            DateTimeOffset value;
            if (!DateTimeOffset.TryParse(
                (string)token,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out value))
            {
                return null;
            }

            // Codex frequently emits these as "all time" sentinel bounds.
            // Treat them as no date filter; otherwise Outlook must Restrict
            // almost every message in every folder for queries like
            // "oldest email", freezing the UI thread on large mailboxes.
            if (!isUpperBound && value <= new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero))
            {
                return null;
            }
            if (isUpperBound && value.Year >= 9999)
            {
                return null;
            }

            return value;
        }
    }
}
