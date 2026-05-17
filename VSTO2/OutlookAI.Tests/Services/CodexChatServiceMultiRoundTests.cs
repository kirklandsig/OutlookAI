using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
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

        /// <summary>
        /// Regression for the production Variants failure 'Missing required
        /// parameter: input[1].call_id'. The real Codex SSE event puts the
        /// cross-reference identifier in 'call_id', not 'id'. The marshaller
        /// must (a) read 'call_id' from the SSE, and (b) emit 'call_id' in
        /// the function_call history item it sends back in round 2.
        /// </summary>
        [Fact]
        public async Task RunTurnAsync_RealCodexShape_UsesCallIdField_RoundTrips()
        {
            var fixt = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            // Real Codex shape: 'call_id' is the cross-reference; 'id' is
            // an internal item id we don't need (left out here entirely).
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"call_id\":\"fc_abc123\",\"name\":\"outlook_count_messages\",\"arguments\":\"{\\\"query\\\":\\\"x\\\"}\"}}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"42\"}\n\n"
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
                    tools.Queue("outlook_count_messages", "{\"count\":42}");

                    var result = await chat.RunTurnAsync(ctx, "count them", tools, sink, CancellationToken.None);

                    Assert.Equal(StopReason.Completed, result.StopReason);
                    // The CallId surfaced to the sink came from 'call_id'.
                    Assert.Equal("fc_abc123", sink.ToolStarts[0].CallId);
                    // Both history items use 'call_id' (not 'id').
                    var fc = ctx.History[1];
                    var fco = ctx.History[2];
                    Assert.Equal("function_call", (string)fc["type"]);
                    Assert.Equal("fc_abc123", (string)fc["call_id"]);
                    Assert.Null((string)fc["id"]);
                    Assert.Equal("function_call_output", (string)fco["type"]);
                    Assert.Equal("fc_abc123", (string)fco["call_id"]);

                    // Verify the round-2 wire body would have passed Codex's
                    // server-side validator (the bit that emitted the
                    // 'Missing required parameter: input[1].call_id' error
                    // pre-fix).
                    var round2 = fake.RequestBodies[1];
                    Assert.Contains("\"call_id\":\"fc_abc123\"", round2);
                    Assert.Contains("\"type\":\"function_call\"", round2);
                    Assert.Contains("\"type\":\"function_call_output\"", round2);
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
                    // Regression: the function_call item we send back in
                    // round 2 must use 'call_id' (not 'id') or the Codex
                    // server returns 'Missing required parameter:
                    // input[N].call_id'.
                    Assert.Equal("call_1", (string)ctx.History[1]["call_id"]);
                    Assert.Null((string)ctx.History[1]["id"]);
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
        public async Task RunTurnAsync_MaxRoundsReached_WhenModelLoopsForever()
        {
            const int maxRounds = 16; // mirror CodexChatService.MaxToolRounds
            var fixt = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            // Queue maxRounds+1 SSE responses, each emitting one function_call.
            // The +1 is insurance; it should never be consumed.
            for (int i = 0; i < maxRounds + 1; i++)
            {
                fake.QueueSse(HttpStatusCode.OK,
                    "data: {\"type\":\"response.output_item.added\",\"item\":{\"type\":\"function_call\",\"id\":\"call_"
                    + i + "\",\"name\":\"outlook_list_folders\",\"arguments\":\"{}\"}}\n\n"
                    + "data: {\"type\":\"response.completed\"}\n\n");
            }
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
                    for (int i = 0; i < maxRounds; i++)
                    {
                        tools.Queue("outlook_list_folders", "{\"folders\":[]}");
                    }

                    var result = await chat.RunTurnAsync(ctx, "loop forever", tools, sink, CancellationToken.None);

                    Assert.Equal(StopReason.MaxRoundsReached, result.StopReason);
                    Assert.Equal(maxRounds, result.RoundsUsed);
                    Assert.Equal(maxRounds, tools.Calls.Count);
                    // The +1 SSE remains queued (only maxRounds requests should have been issued).
                    Assert.Equal(maxRounds, fake.Requests.Count);
                }
            }
            finally
            {
                try { Directory.Delete(fixt.TmpDir, recursive: true); } catch { }
            }
        }

        [Fact]
        public async Task RunTurnAsync_Cancellation_AfterFirstDelta_PreservesPartialText()
        {
            var fixt = MakeAuth();
            var fake = new FakeHttpMessageHandler();
            var gate = new SemaphoreSlim(0, 1);
            // First SSE chunk contains one delta event. After that the stream
            // blocks on `gate` so we can cancel deterministically.
            const string chunk1 =
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"partial \"}\n\n";
            const string chunk2 =
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"never seen\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n";
            fake.QueueRaw(HttpStatusCode.OK, new StreamContent(new PausableStream(chunk1, chunk2, gate))
            {
                Headers = { ContentType = new MediaTypeHeaderValue("text/event-stream") }
            });

            var cts = new CancellationTokenSource();
            var firstDeltaSeen = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var sink = new FirstDeltaSignallingSink(firstDeltaSeen);

            try
            {
                using (fixt.AuthHttp)
                using (fixt.Auth)
                using (var chatHttp = new HttpClient(fake))
                using (var chat = new CodexChatService(fixt.Auth, chatHttp))
                {
                    var ctx = new ConversationContext();
                    var tools = new FakeToolHost();

                    var runTask = chat.RunTurnAsync(ctx, "long answer please", tools, sink, cts.Token);

                    // Wait until the read loop has processed the first delta.
                    await firstDeltaSeen.Task;

                    // Cancel, then release the gate so the stream can finish
                    // draining (allows ReadLineAsync to return so the loop
                    // can observe the cancellation on its next iteration).
                    cts.Cancel();
                    gate.Release();

                    var result = await runTask;

                    Assert.Equal(StopReason.Cancelled, result.StopReason);
                    Assert.Equal("partial ", result.FinalAssistantText);
                    // Partial assistant message was added to history.
                    Assert.Equal(2, ctx.History.Count); // [user, assistant("partial ")]
                    Assert.Equal("assistant", (string)ctx.History[1]["role"]);
                    Assert.Equal("partial ", (string)ctx.History[1]["content"]);
                    Assert.Single(sink.AssistantMessageFinalTexts);
                    Assert.Equal("partial ", sink.AssistantMessageFinalTexts[0]);
                }
            }
            finally
            {
                gate.Dispose();
                try { Directory.Delete(fixt.TmpDir, recursive: true); } catch { }
            }
        }

        // Test helper: signals a TCS as soon as the read loop processes a token delta.
        private sealed class FirstDeltaSignallingSink : CapturingChatEventSink
        {
            private readonly TaskCompletionSource<bool> _firstDelta;
            private int _seen;
            public FirstDeltaSignallingSink(TaskCompletionSource<bool> firstDelta) { _firstDelta = firstDelta; }
            public override void OnTokenDelta(string delta)
            {
                base.OnTokenDelta(delta);
                if (Interlocked.Exchange(ref _seen, 1) == 0)
                {
                    _firstDelta.TrySetResult(true);
                }
            }
        }

        // Test helper: a Stream that yields one chunk, waits on a semaphore,
        // then yields a second chunk. The wait honours cancellation.
        private sealed class PausableStream : Stream
        {
            private readonly byte[] _chunk1;
            private readonly byte[] _chunk2;
            private readonly SemaphoreSlim _gate;
            private int _pos1;
            private int _pos2;
            private bool _passedGate;

            public PausableStream(string chunk1, string chunk2, SemaphoreSlim gate)
            {
                _chunk1 = Encoding.UTF8.GetBytes(chunk1);
                _chunk2 = Encoding.UTF8.GetBytes(chunk2);
                _gate = gate;
            }

            public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken ct)
            {
                if (_pos1 < _chunk1.Length)
                {
                    int n = Math.Min(count, _chunk1.Length - _pos1);
                    Buffer.BlockCopy(_chunk1, _pos1, buffer, offset, n);
                    _pos1 += n;
                    return n;
                }
                if (!_passedGate)
                {
                    await _gate.WaitAsync(ct).ConfigureAwait(false);
                    _passedGate = true;
                }
                if (_pos2 < _chunk2.Length)
                {
                    int n = Math.Min(count, _chunk2.Length - _pos2);
                    Buffer.BlockCopy(_chunk2, _pos2, buffer, offset, n);
                    _pos2 += n;
                    return n;
                }
                return 0;
            }

            public override int Read(byte[] buffer, int offset, int count) =>
                ReadAsync(buffer, offset, count, CancellationToken.None).GetAwaiter().GetResult();
            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }
            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
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
