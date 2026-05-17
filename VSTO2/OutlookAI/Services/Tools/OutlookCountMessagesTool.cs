using System;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_count_messages. Same query syntax as search; returns count only.
    /// </summary>
    public sealed class OutlookCountMessagesTool : IOutlookTool
    {
        public string Name => "outlook_count_messages";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");

            // Phase 3a: every filter is optional, matching outlook_search_messages.
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
                MaxResults      = int.MaxValue,
            };
            var count = surface.CountMessages(search);
            var json = new JObject(new JProperty("count", count));
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }

        private static DateTimeOffset? ParseDate(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            DateTimeOffset value;
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
