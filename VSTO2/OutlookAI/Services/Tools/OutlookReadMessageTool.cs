using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_read_message. Fetches one message by id; body always plaintext.
    /// </summary>
    public sealed class OutlookReadMessageTool : IOutlookTool
    {
        public string Name => "outlook_read_message";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var id = (string)args["message_id"];
            if (string.IsNullOrEmpty(id))
            {
                return Task.FromResult(BuildError("invalid_arguments", "message_id is required"));
            }
            bool includeFullBody = args["include_full_body"]?.Value<bool>() ?? true;

            var detail = surface.ReadMessage(id, includeFullBody);
            if (detail == null)
            {
                return Task.FromResult(BuildError("not_found", "Message " + id + " not found"));
            }

            var json = new JObject(
                new JProperty("id", detail.Id ?? ""),
                new JProperty("subject", detail.Subject ?? ""),
                new JProperty("from", detail.From ?? ""),
                new JProperty("to", new JArray((detail.To ?? new string[0]).Cast<object>())),
                new JProperty("cc", new JArray((detail.Cc ?? new string[0]).Cast<object>())),
                new JProperty("received_at", detail.ReceivedAt.ToString("o")),
                new JProperty("body_plaintext", detail.BodyPlaintext ?? ""),
                new JProperty("body_truncated", detail.BodyTruncated),
                new JProperty("attachments", new JArray(
                    (detail.Attachments ?? new AttachmentSummary[0]).Select(a =>
                        new JObject(
                            new JProperty("filename", a.Filename),
                            new JProperty("size_bytes", a.SizeBytes))))),
                new JProperty("in_reply_to_message_id", detail.InReplyToMessageId),
                new JProperty("conversation_topic", detail.ConversationTopic ?? ""));
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
