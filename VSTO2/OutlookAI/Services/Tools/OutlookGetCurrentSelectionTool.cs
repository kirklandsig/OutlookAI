using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_get_current_selection. Returns the message(s) currently
    /// selected in the user's active Explorer (reading-pane selection).
    /// Read-only. Used by the Inbox Copilot to support "reply to this" /
    /// "summarize this thread" without an upfront search round-trip.
    /// </summary>
    public sealed class OutlookGetCurrentSelectionTool : IOutlookTool
    {
        public string Name => "outlook_get_current_selection";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var includeBodies = (bool?)args["include_full_bodies"] ?? false;
            var maxItemsRaw = (int?)args["max_items"] ?? 5;
            var maxItems = Math.Max(1, Math.Min(20, maxItemsRaw));

            var result = surface.GetCurrentSelection(includeBodies, maxItems);
            var json = ProjectToJson(result, includeBodies);
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }

        private static JObject ProjectToJson(CurrentSelectionResult r, bool includeBodies)
        {
            var arr = new JArray();
            if (r?.Messages != null)
            {
                foreach (var m in r.Messages)
                {
                    var item = new JObject(
                        new JProperty("id", m.Id ?? ""),
                        new JProperty("subject", m.Subject ?? ""),
                        new JProperty("from", m.From ?? ""),
                        new JProperty("received_at", m.ReceivedAt.ToString("o")),
                        new JProperty("conversation_topic", m.ConversationTopic ?? ""),
                        new JProperty("has_attachments", m.Attachments != null && m.Attachments.Count > 0));
                    if (includeBodies)
                    {
                        item.Add("body_plaintext", m.BodyPlaintext ?? "");
                        item.Add("body_truncated", m.BodyTruncated);
                    }
                    else
                    {
                        var snippet = m.BodyPlaintext ?? "";
                        if (snippet.Length > 200) snippet = snippet.Substring(0, 200);
                        item.Add("snippet", snippet);
                    }
                    arr.Add(item);
                }
            }
            return new JObject(
                new JProperty("folder", r?.Folder ?? ""),
                new JProperty("folder_id", r?.FolderId ?? ""),
                new JProperty("count", r?.Count ?? 0),
                new JProperty("messages", arr));
        }
    }
}
