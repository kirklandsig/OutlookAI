using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_set_category. Replaces a message's Categories with the
    /// single given value (matches Outlook UI behavior).
    /// </summary>
    public sealed class OutlookSetCategoryTool : IOutlookTool
    {
        public string Name => "outlook_set_category";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var id = (string)args["message_id"];
            var category = (string)args["category"];
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(category))
            {
                return Task.FromResult(BuildError("invalid_arguments",
                    "message_id and category are required"));
            }
            surface.SetCategory(id, category);
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
