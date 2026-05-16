using Newtonsoft.Json.Linq;
using OutlookAI.Services.Chat;
using Xunit;

namespace OutlookAI.Tests.Services.Chat
{
    public class ConversationStoreTests
    {
        [Fact]
        public void Append_AndSnapshot_RoundTrips()
        {
            var s = new ConversationStore();
            s.Append(JObject.Parse("{\"type\":\"message\",\"role\":\"user\",\"content\":\"hi\"}"));
            var snap = s.Snapshot();
            Assert.Single(snap);
            Assert.Equal("hi", (string)snap[0]["content"]);
        }

        [Fact]
        public void Clear_RemovesAll()
        {
            var s = new ConversationStore();
            s.Append(new JObject(new JProperty("type", "message")));
            s.Clear();
            Assert.Empty(s.Snapshot());
        }

        [Fact]
        public void ExportForClipboard_RendersUserAndAssistantMessages()
        {
            var s = new ConversationStore();
            s.Append(JObject.Parse("{\"type\":\"message\",\"role\":\"user\",\"content\":\"hi\"}"));
            s.Append(JObject.Parse("{\"type\":\"message\",\"role\":\"assistant\",\"content\":\"hello\"}"));
            var text = s.ExportForClipboard();
            Assert.Contains("[user]", text);
            Assert.Contains("hi", text);
            Assert.Contains("[assistant]", text);
            Assert.Contains("hello", text);
        }

        [Fact]
        public void ExportForClipboard_RendersFunctionCallsAndOutputs()
        {
            var s = new ConversationStore();
            s.Append(JObject.Parse(
                "{\"type\":\"function_call\",\"name\":\"outlook_get_current_compose_state\",\"arguments\":\"{}\"}"));
            s.Append(JObject.Parse(
                "{\"type\":\"function_call_output\",\"output\":\"{\\\"subject\\\":\\\"x\\\"}\"}"));
            var text = s.ExportForClipboard();
            Assert.Contains("[tool call]", text);
            Assert.Contains("outlook_get_current_compose_state", text);
            Assert.Contains("[tool result]", text);
        }

        [Fact]
        public void TwoStores_AreIsolated()
        {
            var a = new ConversationStore();
            var b = new ConversationStore();
            a.Append(new JObject(new JProperty("type", "message"), new JProperty("role", "user"),
                new JProperty("content", "in-a")));
            Assert.Single(a.Snapshot());
            Assert.Empty(b.Snapshot());
        }
    }
}
