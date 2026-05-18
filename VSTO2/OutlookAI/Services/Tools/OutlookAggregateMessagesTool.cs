using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_aggregate_messages. Group messages by sender, day,
    /// or folder and return the top-N buckets by count. Used by stats
    /// and out-of-office-catchup reports.
    /// </summary>
    public sealed class OutlookAggregateMessagesTool : IOutlookTool
    {
        public string Name => "outlook_aggregate_messages";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            try
            {
                ct.ThrowIfCancellationRequested();
                var args = AggregateMessagesArgsParser.Parse(argsJson);
                var buckets = surface.AggregateMessages(args, ct) ?? new AggregationBucket[0];

                var total = buckets.Sum(b => b.Count);

                var json = new JObject(
                    new JProperty("buckets", new JArray(buckets.Select(b =>
                        new JObject(
                            new JProperty("label", b.Label ?? ""),
                            new JProperty("count", b.Count))))),
                    new JProperty("total", total));
                return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
            }
            catch (OperationCanceledException)
            {
                return Task.FromResult(BuildError("cancelled", "Aggregation cancelled by user."));
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
