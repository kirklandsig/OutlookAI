using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_search_messages. Searches via DASL Restrict and returns
    /// up to <c>max_results</c> hits (default 25, hard cap 100).
    /// </summary>
    public sealed class OutlookSearchMessagesTool : IOutlookTool
    {
        public string Name => "outlook_search_messages";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");

            // Phase 3a: every filter is optional. The model can call with
            // zero args to enumerate the inbox (capped by max_results).
            // Previous schema required 'query' - that constraint is gone.
            int maxResults = args["max_results"]?.Value<int>() ?? 25;
            if (maxResults < 1) maxResults = 1;
            if (maxResults > 100) maxResults = 100;

            var search = new SearchMessagesArgs
            {
                Query           = (string)args["query"],
                From            = (string)args["from"],
                SubjectContains = (string)args["subject_contains"],
                BodyContains    = (string)args["body_contains"],
                HasAttachment   = args["has_attachment"] != null ? (bool?)(bool)args["has_attachment"] : null,
                IsUnread        = args["is_unread"]      != null ? (bool?)(bool)args["is_unread"]      : null,
                IsFlagged       = args["is_flagged"]     != null ? (bool?)(bool)args["is_flagged"]     : null,
                Importance      = (string)args["importance"],
                FolderId        = (string)args["folder_id"],
                DateFrom        = ParseDate(args["date_from"]),
                DateTo          = ParseDate(args["date_to"]),
                MaxResults      = maxResults,
            };

            var hits = surface.SearchMessages(search) ?? new MessageSummary[0];
            var json = new JObject(
                new JProperty("messages", new JArray(hits.Select(m =>
                    new JObject(
                        new JProperty("id", m.Id ?? ""),
                        new JProperty("subject", m.Subject ?? ""),
                        new JProperty("from", m.From ?? ""),
                        new JProperty("to", new JArray((m.To ?? new string[0]).Cast<object>())),
                        new JProperty("received_at", m.ReceivedAt.ToString("o")),
                        new JProperty("snippet", m.Snippet ?? ""),
                        new JProperty("has_attachments", m.HasAttachments))))));
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }

        private static DateTimeOffset? ParseDate(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            DateTimeOffset value;
            // AssumeUniversal + AdjustToUniversal: parse the ISO-8601
            // value at face value (UTC if it ends in Z, with declared offset
            // otherwise) and surface it as UTC. The Codex backend sends
            // dates in UTC; we never want a localtime offset round-trip.
            return DateTimeOffset.TryParse(
                (string)token,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out value) ? value : (DateTimeOffset?)null;
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
