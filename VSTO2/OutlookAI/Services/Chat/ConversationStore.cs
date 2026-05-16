using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Chat
{
    /// <summary>
    /// Per-Inspector in-memory chat state. Snapshot is thread-safe; mutation
    /// happens under a lock so the WebView2 push and the chat service can
    /// touch this from different threads.
    /// </summary>
    public sealed class ConversationStore
    {
        private readonly object _lock = new object();
        private readonly List<JObject> _history = new List<JObject>();

        public IReadOnlyList<JObject> Snapshot()
        {
            lock (_lock) return _history.ToArray();
        }

        public void Append(JObject item)
        {
            lock (_lock) _history.Add(item);
        }

        public void AppendRange(IEnumerable<JObject> items)
        {
            lock (_lock) _history.AddRange(items);
        }

        public void Clear()
        {
            lock (_lock) _history.Clear();
        }

        public int Count
        {
            get { lock (_lock) return _history.Count; }
        }

        public string ExportForClipboard()
        {
            var sb = new StringBuilder();
            lock (_lock)
            {
                foreach (var item in _history)
                {
                    var type = (string)item["type"] ?? "?";
                    if (type == "message")
                    {
                        var role = (string)item["role"];
                        var content = ExtractText(item["content"]);
                        sb.AppendLine("[" + role + "]");
                        sb.AppendLine(content);
                        sb.AppendLine();
                    }
                    else if (type == "function_call")
                    {
                        sb.AppendLine("[tool call] " + item["name"] + " " + item["arguments"]);
                    }
                    else if (type == "function_call_output")
                    {
                        sb.AppendLine("[tool result] " + item["output"]);
                        sb.AppendLine();
                    }
                }
            }
            return sb.ToString();
        }

        private static string ExtractText(JToken token)
        {
            if (token == null) return "";
            if (token.Type == JTokenType.String) return (string)token;
            if (token.Type == JTokenType.Array)
            {
                var sb = new StringBuilder();
                foreach (var part in (JArray)token)
                {
                    var text = (string)part["text"];
                    if (!string.IsNullOrEmpty(text)) sb.Append(text);
                }
                return sb.ToString();
            }
            return token.ToString();
        }
    }
}
