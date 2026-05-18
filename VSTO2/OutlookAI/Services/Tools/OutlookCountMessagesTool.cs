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
            try
            {
                ct.ThrowIfCancellationRequested();
                var search = SearchMessagesArgsParser.ParseCount(argsJson);
                var count = surface.CountMessages(search, ct);
                var json = new JObject(new JProperty("count", count));
                return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Search cancelled by user."));
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
