using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class CodexChatServiceMultiRoundTests
    {
        private static (CodexAuthService Auth, HttpClient AuthHttp, string TmpDir) MakeAuth()
        {
            var tmp = Path.Combine(Path.GetTempPath(), "outlookai-mr", Path.GetRandomFileName());
            Directory.CreateDirectory(tmp);
            var path = Path.Combine(tmp, "auth.json");
            File.WriteAllText(path,
                "{\"tokens\":{\"access_token\":\"sk-test\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");
            var http = new HttpClient(new FakeHttpMessageHandler());
            return (new CodexAuthService(path, http), http, tmp);
        }

        [Fact]
        public async Task RunTurnAsync_SingleRound_NoToolCalls_ReturnsCompleted()
        {
            var fixt = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"hello\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            try
            {
                using (fixt.AuthHttp)
                using (fixt.Auth)
                using (var chatHttp = new HttpClient(fake))
                using (var chat = new CodexChatService(fixt.Auth, chatHttp))
                {
                    var ctx = new ConversationContext { SystemInstructions = "Be brief." };
                    var sink = new CapturingChatEventSink();
                    var tools = new FakeToolHost();

                    var result = await chat.RunTurnAsync(ctx, "say hi", tools, sink, CancellationToken.None);

                    Assert.Equal(StopReason.Completed, result.StopReason);
                    Assert.Equal("hello", result.FinalAssistantText);
                    Assert.Equal(1, result.RoundsUsed);
                    Assert.Equal("hello", sink.StreamedText.ToString());
                    Assert.Single(sink.AssistantMessageFinalTexts);
                    Assert.Empty(tools.Calls);
                    // Single round: 1 user item + 1 assistant message in history; appended duplicates assistant.
                    Assert.Equal(2, ctx.History.Count);
                    Assert.Single(result.AppendedItems);
                    Assert.Equal(1, sink.RoundBoundaries);
                    // Request shape verified by snapshotting the outgoing body.
                    Assert.Single(fake.Requests);
                    Assert.Equal("Bearer sk-test", fake.Requests[0].Headers.Authorization.ToString());
                    Assert.Contains("\"model\":\"gpt-5.5\"", fake.RequestBodies[0]);
                    Assert.Contains("\"parallel_tool_calls\":true", fake.RequestBodies[0]);
                    Assert.Contains("\"stream\":true", fake.RequestBodies[0]);
                    Assert.Contains("\"store\":false", fake.RequestBodies[0]);
                    Assert.Contains("\"tools\":[", fake.RequestBodies[0]);
                    Assert.Contains("outlook_get_current_compose_state", fake.RequestBodies[0]);
                }
            }
            finally
            {
                try { Directory.Delete(fixt.TmpDir, recursive: true); } catch { }
            }
        }
    }
}
