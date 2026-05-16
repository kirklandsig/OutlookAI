using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_list_recent_threads_with. Groups Inbox+Sent conversations
    /// by ConversationID for one recipient.
    /// </summary>
    public sealed class OutlookListRecentThreadsWithTool : IOutlookTool
    {
        public string Name => "outlook_list_recent_threads_with";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var recipient = (string)args["recipient_email"];
            if (string.IsNullOrWhiteSpace(recipient))
            {
                return Task.FromResult(BuildError("invalid_arguments", "recipient_email is required"));
            }
            int maxThreads = args["max_threads"]?.Value<int>() ?? 5;
            if (maxThreads < 1) maxThreads = 1;
            if (maxThreads > 20) maxThreads = 20;

            var threads = surface.ListRecentThreadsWith(recipient, maxThreads) ?? new ThreadSummary[0];
            var json = new JObject(
                new JProperty("threads", new JArray(threads.Select(t =>
                    new JObject(
                        new JProperty("thread_topic", t.ThreadTopic ?? ""),
                        new JProperty("last_message_at", t.LastMessageAt.ToString("o")),
                        new JProperty("message_count", t.MessageCount),
                        new JProperty("snippet", t.Snippet ?? ""),
                        new JProperty("thread_id", t.ThreadId ?? ""))))));
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
