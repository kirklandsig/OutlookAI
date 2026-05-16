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

        [Fact]
        public async Task RunTurnAsync_ToolCall_DispatchesAndAppendsFunctionOutput_ThenCompletes()
        {
            var fixt = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            // Round 1: model emits a function_call for outlook_get_current_compose_state.
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"call_1\",\"name\":\"outlook_get_current_compose_state\",\"arguments\":\"{}\"}}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            // Round 2: model produces the final assistant text given the tool's reply.
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"Subject was X\"}\n\n"
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
                    tools.Queue("outlook_get_current_compose_state", "{\"subject\":\"X\"}");

                    var result = await chat.RunTurnAsync(ctx, "what's the subject", tools, sink, CancellationToken.None);

                    Assert.Equal(StopReason.Completed, result.StopReason);
                    Assert.Equal(2, result.RoundsUsed);
                    Assert.Equal("Subject was X", result.FinalAssistantText);
                    Assert.Single(tools.Calls);
                    Assert.Equal("outlook_get_current_compose_state", tools.Calls[0].Name);
                    Assert.Equal("{}", tools.Calls[0].ArgsJson);
                    Assert.Equal(2, sink.RoundBoundaries);
                    Assert.Single(sink.ToolStarts);
                    Assert.Equal("call_1", sink.ToolStarts[0].CallId);
                    Assert.Single(sink.ToolResults);
                    Assert.True(sink.ToolResults[0].Ok);
                    // History after turn:
                    //   [0] user "what's the subject"
                    //   [1] function_call (call_1)
                    //   [2] function_call_output (call_1 -> {"subject":"X"})
                    //   [3] assistant "Subject was X"
                    Assert.Equal(4, ctx.History.Count);
                    Assert.Equal("function_call", (string)ctx.History[1]["type"]);
                    Assert.Equal("function_call_output", (string)ctx.History[2]["type"]);
                    Assert.Equal("call_1", (string)ctx.History[2]["call_id"]);
                    Assert.Equal("{\"subject\":\"X\"}", (string)ctx.History[2]["output"]);
                    Assert.Equal(2, fake.Requests.Count);
                }
            }
            finally
            {
                try { Directory.Delete(fixt.TmpDir, recursive: true); } catch { }
            }
        }

        [Fact]
        public async Task RunTurnAsync_ParallelToolCalls_AreDispatchedConcurrently_AndBothLogged()
        {
            var fixt = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            // Round 1: TWO function_call items in one SSE stream.
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"call_a\",\"name\":\"outlook_list_folders\",\"arguments\":\"{}\"}}\n\n"
                + "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"call_b\",\"name\":\"outlook_count_messages\",\"arguments\":\"{\\\"query\\\":\\\"x\\\"}\"}}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            // Round 2: assistant final answer.
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"done\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            try
            {
                using (fixt.AuthHttp)
                using (fixt.Auth)
                using (var chatHttp = new HttpClient(fake))
                using (var chat = new CodexChatService(fixt.Auth, chatHttp))
                {
                    var ctx = new ConversationContext();
                    var sink = new CapturingChatEventSink();
                    var tools = new FakeToolHost();
                    tools.Queue("outlook_list_folders", "{\"folders\":[]}");
                    tools.Queue("outlook_count_messages", "{\"count\":3}");

                    var result = await chat.RunTurnAsync(ctx, "do both", tools, sink, CancellationToken.None);

                    Assert.Equal(StopReason.Completed, result.StopReason);
                    Assert.Equal(2, result.RoundsUsed);
                    Assert.Equal(2, tools.Calls.Count);
                    Assert.Equal(2, sink.ToolStarts.Count);
                    Assert.Equal(2, sink.ToolResults.Count);
                    // Both function_call + function_call_output pairs landed in history.
                    var fcalls = ctx.History.FindAll(it => (string)it["type"] == "function_call");
                    var fouts  = ctx.History.FindAll(it => (string)it["type"] == "function_call_output");
                    Assert.Equal(2, fcalls.Count);
                    Assert.Equal(2, fouts.Count);
                }
            }
            finally
            {
                try { Directory.Delete(fixt.TmpDir, recursive: true); } catch { }
            }
        }

        [Fact]
        public async Task RunTurnAsync_ToolThrows_ProducesErrorEnvelope_AndContinuesNextRound()
        {
            var fixt = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"call_x\",\"name\":\"outlook_read_message\",\"arguments\":\"{\\\"message_id\\\":\\\"abc\\\"}\"}}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"could not read\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            try
            {
                using (fixt.AuthHttp)
                using (fixt.Auth)
                using (var chatHttp = new HttpClient(fake))
                using (var chat = new CodexChatService(fixt.Auth, chatHttp))
                {
                    var ctx = new ConversationContext();
                    var sink = new CapturingChatEventSink();
                    var tools = new FakeToolHost();
                    tools.QueueThrow("outlook_read_message", new InvalidOperationException("boom"));

                    var result = await chat.RunTurnAsync(ctx, "read it", tools, sink, CancellationToken.None);

                    Assert.Equal(StopReason.Completed, result.StopReason);
                    Assert.Single(sink.ToolResults);
                    Assert.False(sink.ToolResults[0].Ok);
                    // function_call_output should carry an error envelope.
                    var fout = ctx.History.Find(it => (string)it["type"] == "function_call_output");
                    Assert.NotNull(fout);
                    Assert.Contains("\"error\"", (string)fout["output"]);
                }
            }
            finally
            {
                try { Directory.Delete(fixt.TmpDir, recursive: true); } catch { }
            }
        }
    }
}
