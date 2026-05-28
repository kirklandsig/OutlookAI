using System.Collections.Generic;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Result of a message search: the clamped page of summaries the model
    /// sees, plus the pre-clamp total and whether the page was truncated.
    /// TotalMatches is exact where the collection path saw every match (the
    /// common from:/subject:/body: case) and a floor where an early-stop
    /// fired (to:/broad-no-filter scans); in the floor case Truncated is
    /// forced true so the model never believes a capped page is complete.
    /// </summary>
    public sealed class SearchResult
    {
        public IReadOnlyList<MessageSummary> Messages { get; set; } = new MessageSummary[0];
        public int TotalMatches { get; set; }
        public bool Truncated { get; set; }
    }
}
