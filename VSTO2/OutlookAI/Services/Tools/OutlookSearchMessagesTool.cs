using System;
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
            var query = (string)args["query"];
            if (string.IsNullOrWhiteSpace(query))
            {
                return Task.FromResult(BuildError("invalid_arguments", "query is required"));
            }

            int maxResults = args["max_results"]?.Value<int>() ?? 25;
            if (maxResults < 1) maxResults = 1;
            if (maxResults > 100) maxResults = 100;

            var search = new SearchMessagesArgs
            {
                Query = query,
                FolderId = (string)args["folder_id"],
                DateFrom = ParseDate(args["date_from"]),
                DateTo = ParseDate(args["date_to"]),
                MaxResults = maxResults,
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
            return DateTimeOffset.TryParse((string)token, out value) ? value : (DateTimeOffset?)null;
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
