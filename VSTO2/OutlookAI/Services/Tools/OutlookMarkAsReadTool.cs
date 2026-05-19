using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_mark_as_read. Sets or clears the UnRead flag on a message.
    /// </summary>
    public sealed class OutlookMarkAsReadTool : IOutlookTool
    {
        public string Name => "outlook_mark_as_read";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var id = (string)args["message_id"];
            var readToken = args["read"];
            if (string.IsNullOrEmpty(id) || readToken == null || readToken.Type != JTokenType.Boolean)
            {
                return Task.FromResult(BuildError("invalid_arguments",
                    "message_id and read (boolean) are required"));
            }
            surface.MarkAsRead(id, (bool)readToken);
            return Task.FromResult("{\"ok\":true}");
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
