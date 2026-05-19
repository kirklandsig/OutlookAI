using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Chat
{
    /// <summary>
    /// Per-turn context for <c>CodexChatService.RunTurnAsync</c>. Carries the
    /// accumulated conversation history, the system instructions to seed the
    /// turn, the (nullable) reasoning-effort override, and which tools the
    /// chat service should expose for this turn.
    /// </summary>
    public sealed class ConversationContext
    {
        public string SystemInstructions { get; set; } = "";
        public List<JObject> History { get; set; } = new List<JObject>();
        public string ReasoningEffortOverride { get; set; }   // null => Config.ReasoningEffort
        public bool IncludeWriteTools { get; set; } = true;
    }
}
