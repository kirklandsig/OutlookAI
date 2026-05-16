using System.Collections.Generic;
using System.Text;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Records every event a chat turn emits, for assertions in chat-service
    /// tests. Replaces the production WebView2-backed sink.
    /// </summary>
    public sealed class CapturingChatEventSink : OutlookAI.Services.Chat.ChatEventSink
    {
        public StringBuilder StreamedText { get; } = new StringBuilder();
        public List<(string CallId, string Name, string ArgsJson)> ToolStarts { get; }
            = new List<(string, string, string)>();
        public List<(string CallId, bool Ok, string Summary, string ResultJson)> ToolResults { get; }
            = new List<(string, bool, string, string)>();
        public List<string> AssistantMessageFinalTexts { get; } = new List<string>();
        public List<string> Errors { get; } = new List<string>();
        public int RoundBoundaries { get; private set; }

        public override void OnTokenDelta(string delta) => StreamedText.Append(delta);
        public override void OnToolCallStart(string callId, string name, string argsJson)
            => ToolStarts.Add((callId, name, argsJson));
        public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            => ToolResults.Add((callId, ok, summary, resultJson));
        public override void OnAssistantMessageComplete(string text)
            => AssistantMessageFinalTexts.Add(text);
        public override void OnError(string message) => Errors.Add(message);
        public override void OnRoundBoundary() => RoundBoundaries++;
    }
}
