using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_get_current_compose_state.
    /// Reads the compose-window state for the Inspector that owns this chat.
    /// </summary>
    public sealed class OutlookGetCurrentComposeStateTool : IOutlookTool
    {
        public string Name => "outlook_get_current_compose_state";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            bool fullBody = args["include_full_body"]?.Value<bool>() ?? false;

            var state = surface.GetCurrentComposeState(fullBody);
            var json = new JObject(
                new JProperty("subject", state.Subject ?? ""),
                new JProperty("recipients", new JObject(
                    new JProperty("to", new JArray((state.ToRecipients ?? new string[0]).Cast<object>())),
                    new JProperty("cc", new JArray((state.CcRecipients ?? new string[0]).Cast<object>())),
                    new JProperty("bcc", new JArray((state.BccRecipients ?? new string[0]).Cast<object>())))),
                new JProperty("sender_name", state.SenderName ?? ""),
                new JProperty("sender_email", state.SenderEmail ?? ""),
                new JProperty("body_plaintext", state.BodyPlaintext ?? ""),
                new JProperty("body_truncated", state.BodyTruncated),
                new JProperty("attachments", new JArray(
                    (state.Attachments ?? new AttachmentSummary[0]).Select(a =>
                        new JObject(
                            new JProperty("filename", a.Filename),
                            new JProperty("size_bytes", a.SizeBytes))))));
            if (state.InReplyTo != null)
            {
                json["in_reply_to"] = new JObject(
                    new JProperty("thread_topic", state.InReplyTo.ThreadTopic ?? ""),
                    new JProperty("last_n_messages", new JArray(
                        (state.InReplyTo.LastNMessages ?? new ThreadMessage[0]).Select(m =>
                            new JObject(
                                new JProperty("from", m.From ?? ""),
                                new JProperty("received_at", m.ReceivedAt.ToString("o")),
                                new JProperty("snippet", m.Snippet ?? ""))))));
            }
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }
    }
}
