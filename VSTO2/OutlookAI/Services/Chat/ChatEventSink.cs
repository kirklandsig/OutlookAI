namespace OutlookAI.Services.Chat
{
    /// <summary>
    /// Streaming callback surface for one chat turn. Override only the
    /// callbacks you care about. Defaults are no-ops so callers can construct
    /// a bare <c>ChatEventSink</c> when they don't care about events
    /// (e.g. <c>CodexChatService.RunTurnAsync</c> uses an empty instance as
    /// its default sink).
    /// </summary>
    public class ChatEventSink
    {
        public virtual void OnTokenDelta(string delta) { }
        public virtual void OnToolCallStart(string callId, string name, string argsJson) { }
        public virtual void OnToolCallResult(string callId, bool ok, string summary, string resultJson) { }
        public virtual void OnAssistantMessageComplete(string text) { }
        public virtual void OnError(string message) { }
        public virtual void OnRoundBoundary() { }
    }
}
