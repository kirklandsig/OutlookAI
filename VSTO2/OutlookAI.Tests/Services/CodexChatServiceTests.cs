using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    public class CodexChatServiceTests : IDisposable
    {
        private readonly string _tmpDir;
        private readonly string _authPath;

        public CodexChatServiceTests()
        {
            _tmpDir = Path.Combine(Path.GetTempPath(), "outlookai-chat", Path.GetRandomFileName());
            Directory.CreateDirectory(_tmpDir);
            _authPath = Path.Combine(_tmpDir, "auth.json");
            // Seed a fresh token so ProcessEmailAsync doesn't try to refresh.
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk-test\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");
        }

        public void Dispose()
        {
            try { Directory.Delete(_tmpDir, recursive: true); } catch { }
        }

        [Fact]
        public async Task ProcessEmailAsync_SendsResponsesRequestWithBearer()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"Hello\"}\n\n"
                + "data: {\"type\":\"response.output_text.delta\",\"delta\":\" world\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var result = await chat.ProcessEmailAsync(
                    CodexChatService.ActionType.Proofread, "helo world");

                Assert.Equal("Hello world", result);
                Assert.Single(fake.Requests);
                Assert.Equal("https://chatgpt.com/backend-api/codex/responses",
                    fake.Requests[0].RequestUri.ToString());
                Assert.Equal("Bearer sk-test",
                    fake.Requests[0].Headers.Authorization.ToString());
                Assert.Contains("\"model\":\"gpt-5.5\"", fake.RequestBodies[0]);
                Assert.Contains("\"stream\":true", fake.RequestBodies[0]);
                Assert.Contains("\"store\":false", fake.RequestBodies[0]);
            }
        }

        [Fact]
        public async Task ProcessEmailAsync_PrefersOutputTextDoneOverAccumulatedDeltas()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.delta\",\"delta\":\"draft \"}\n\n"
                + "data: {\"type\":\"response.output_text.done\",\"text\":\"FINAL\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var result = await chat.ProcessEmailAsync(
                    CodexChatService.ActionType.Revise, "any");

                Assert.Equal("FINAL", result);
            }
        }

        [Fact]
        public async Task ProcessEmailAsync_ThrowsWithBackendErrorBody_OnNonSuccess()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueText((HttpStatusCode)429, // TooManyRequests
                "{\"detail\":\"rate limited\"}");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var ex = await Assert.ThrowsAsync<InvalidOperationException>(
                    () => chat.ProcessEmailAsync(
                        CodexChatService.ActionType.Proofread, "x"));

                Assert.Contains("ChatGPT Codex backend error", ex.Message);
                Assert.Contains("rate limited", ex.Message);
            }
        }
    }
}
