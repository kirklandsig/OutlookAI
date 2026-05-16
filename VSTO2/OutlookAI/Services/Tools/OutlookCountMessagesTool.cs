using System;
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
            var query = (string)args["query"];
            if (string.IsNullOrWhiteSpace(query))
            {
                return Task.FromResult(BuildError("invalid_arguments", "query is required"));
            }

            var search = new SearchMessagesArgs
            {
                Query = query,
                FolderId = (string)args["folder_id"],
                DateFrom = ParseDate(args["date_from"]),
                DateTo = ParseDate(args["date_to"]),
                MaxResults = int.MaxValue,
            };
            var count = surface.CountMessages(search);
            var json = new JObject(new JProperty("count", count));
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
