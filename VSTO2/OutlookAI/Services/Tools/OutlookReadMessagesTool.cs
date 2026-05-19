using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_read_messages. Bulk-read message details by short
    /// ID array. Used by reports that need bodies for many messages
    /// (action items, topic status, conversation summaries). Replaces
    /// many outlook_read_message round trips with one call.
    /// </summary>
    public sealed class OutlookReadMessagesTool : IOutlookTool
    {
        private const int DefaultMaxItems = 25;
        private const int MaxItemsCap = 100;

        public string Name => "outlook_read_messages";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);

                string[] ids = (args["ids"] as JArray)?
                    .Select(t => (string)t)
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray() ?? new string[0];

                bool includeBody = args["include_body"]?.Type == JTokenType.Boolean
                    ? args["include_body"].Value<bool>()
                    : true;

                int maxItems = args["max_items"]?.Value<int>() ?? DefaultMaxItems;
                if (maxItems < 1) maxItems = 1;
                if (maxItems > MaxItemsCap) maxItems = MaxItemsCap;

                IReadOnlyList<MessageDetail> messages;
                if (ids.Length == 0)
                {
                    messages = new MessageDetail[0];
                }
                else
                {
                    messages = surface.ReadMessages(ids, includeBody, maxItems, ct) ?? new MessageDetail[0];
                }

                var json = new JObject(
                    new JProperty("messages", new JArray(messages.Select(m =>
                        new JObject(
                            new JProperty("id", m.Id ?? ""),
                            new JProperty("subject", m.Subject ?? ""),
                            new JProperty("from", m.From ?? ""),
                            new JProperty("to", new JArray((m.To ?? new string[0]).Cast<object>())),
                            new JProperty("cc", new JArray((m.Cc ?? new string[0]).Cast<object>())),
                            new JProperty("received_at", m.ReceivedAt.ToString("o")),
                            new JProperty("body_plaintext", m.BodyPlaintext ?? ""),
                            new JProperty("body_truncated", m.BodyTruncated),
                            new JProperty("attachments", new JArray((m.Attachments ?? new AttachmentSummary[0]).Select(a =>
                                new JObject(
                                    new JProperty("filename", a.Filename ?? ""),
                                    new JProperty("size_bytes", a.SizeBytes))))),
                            new JProperty("in_reply_to_message_id", m.InReplyToMessageId ?? ""),
                            new JProperty("conversation_topic", m.ConversationTopic ?? ""))))));
                return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Read cancelled by user."));
            }
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
