using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_create_draft. Creates a draft in the Drafts folder; never sends.
    /// </summary>
    public sealed class OutlookCreateDraftTool : IOutlookTool
    {
        public string Name => "outlook_create_draft";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var subject = (string)args["subject"];
            var body = (string)args["body_plaintext"];
            if (string.IsNullOrEmpty(subject) || body == null)
            {
                return Task.FromResult(BuildError("invalid_arguments",
                    "subject and body_plaintext are required"));
            }

            var created = surface.CreateDraft(new CreateDraftArgs
            {
                Subject = subject,
                BodyPlaintext = body,
                To = ToStringList(args["to"]),
                Cc = ToStringList(args["cc"]),
                InReplyToMessageId = (string)args["in_reply_to_message_id"],
            });

            var json = new JObject(
                new JProperty("draft_id", created.DraftId ?? ""),
                new JProperty("location", created.Location ?? "Drafts"));
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }

        private static IReadOnlyList<string> ToStringList(JToken token)
        {
            if (token == null || token.Type != JTokenType.Array) return null;
            return ((JArray)token).Select(t => (string)t).Where(s => !string.IsNullOrEmpty(s)).ToList();
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
