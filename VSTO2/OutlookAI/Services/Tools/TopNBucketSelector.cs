using System.Collections.Generic;
using System.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Sorts AggregationBucket records by count descending (alphabetical
    /// label tiebreak for deterministic output across runs) and returns
    /// the first N. Defensive against null / empty / non-positive N.
    /// </summary>
    public static class TopNBucketSelector
    {
        public static IReadOnlyList<AggregationBucket> TakeTop(
            IEnumerable<AggregationBucket> buckets, int n)
        {
            if (buckets == null || n <= 0) return new AggregationBucket[0];
            return buckets
                .OrderByDescending(b => b.Count)
                .ThenBy(b => b.Label, System.StringComparer.OrdinalIgnoreCase)
                .Take(n)
                .ToList();
        }
    }
}
