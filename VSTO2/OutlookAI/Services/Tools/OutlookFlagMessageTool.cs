using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_flag_message. Sets follow-up flag status: none | todo | complete.
    /// </summary>
    public sealed class OutlookFlagMessageTool : IOutlookTool
    {
        public string Name => "outlook_flag_message";

        private static readonly string[] ValidFlags = { "none", "todo", "complete" };

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var id = (string)args["message_id"];
            var flag = (string)args["flag"];
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(flag) ||
                Array.IndexOf(ValidFlags, flag) < 0)
            {
                return Task.FromResult(BuildError("invalid_arguments",
                    "message_id and flag (none|todo|complete) are required"));
            }
            surface.FlagMessage(id, flag);
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
