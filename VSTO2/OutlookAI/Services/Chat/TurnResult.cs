using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Chat
{
    /// <summary>
    /// Outcome of a single <c>RunTurnAsync</c> invocation. The history items
    /// appended to <see cref="ConversationContext.History"/> during this turn
    /// are duplicated here for callers that want to re-render or log them
    /// without re-walking the entire history.
    /// </summary>
    public sealed class TurnResult
    {
        public StopReason StopReason { get; set; }
        public string FinalAssistantText { get; set; } = "";
        public int RoundsUsed { get; set; }
        public IReadOnlyList<JObject> AppendedItems { get; set; } = new List<JObject>();
        public string ErrorMessage { get; set; }
    }
}
