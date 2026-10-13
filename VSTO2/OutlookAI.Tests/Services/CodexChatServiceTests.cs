using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    [Collection("Config")]
    public class CodexChatServiceTests : IDisposable
    {
        private readonly string _tmpDir;
        private readonly string _authPath;
        private readonly ConfigStateScope _scope = new ConfigStateScope();

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
            _scope.Dispose();
            try { Directory.Delete(_tmpDir, recursive: true); } catch { }
        }

        [Fact]
        public async Task ProcessEmailAsync_SendsResponsesRequestWithBearer()
        {
            Config.Model = "gpt-5.5";
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
                var ex = await Assert.ThrowsAsync<CodexBackendException>(
                    () => chat.ProcessEmailAsync(
                        CodexChatService.ActionType.Proofread, "x"));

                Assert.Equal(429, (int)ex.StatusCode);
                Assert.Contains("ChatGPT Codex backend error", ex.Message);
                Assert.Contains("rate limited", ex.Message);
            }
        }

        private static IReadOnlyList<JObject> TextAndImage() => new List<JObject>
        {
            new JObject(new JProperty("type", "input_text"), new JProperty("text", "Plan the cleanup of this image.")),
            new JObject(new JProperty("type", "input_image"), new JProperty("image_url", "data:image/jpeg;base64,AAAA")),
        };

        private static JObject Schema() => JObject.Parse("{\"type\":\"object\",\"additionalProperties\":false,\"required\":[\"ok\"],\"properties\":{\"ok\":{\"type\":\"boolean\"}}}");

        [Fact]
        public async Task CompleteStructuredAsync_SendsTheImageAndAStrictSchema()
        {
            Config.Model = "gpt-5.6-sol";
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"type\":\"response.output_text.done\",\"text\":\"{\\\"ok\\\":true}\"}\n\n"
                + "data: {\"type\":\"response.completed\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var text = await chat.CompleteStructuredAsync("instructions here", TextAndImage(), "page_plan", Schema(), "Low", CancellationToken.None);

                Assert.Equal("{\"ok\":true}", text);
                var body = JObject.Parse(fake.RequestBodies[0]);
                Assert.Equal("gpt-5.6-sol", (string)body["model"]);
                Assert.Equal("instructions here", (string)body["instructions"]);
                var content = (JArray)body["input"][0]["content"];
                Assert.Equal("input_image", (string)content[1]["type"]);
                Assert.StartsWith("data:image/jpeg;base64,", (string)content[1]["image_url"]);
                var format = body["text"]["format"];
                Assert.Equal("json_schema", (string)format["type"]);
                Assert.Equal("page_plan", (string)format["name"]);
                Assert.True((bool)format["strict"]);
                Assert.True(JToken.DeepEquals(Schema(), format["schema"]));
                Assert.Equal("low", (string)body["reasoning"]["effort"]);
                Assert.False((bool)body["store"]);
                Assert.True((bool)body["stream"]);
                Assert.Null(body["tools"]);
            }
        }

        [Fact]
        public async Task CompleteStructuredAsync_OmitsAnEffortTheModelDoesntTake()
        {
            Config.Model = "gpt-5.5";
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK, "data: {\"type\":\"response.output_text.done\",\"text\":\"{}\"}\n\n");
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                await chat.CompleteStructuredAsync("i", TextAndImage(), "page_plan", Schema(), "Ultra", CancellationToken.None);

                Assert.Equal(JTokenType.Null, JObject.Parse(fake.RequestBodies[0])["reasoning"].Type);
            }
        }

        [Fact]
        public async Task CompleteStructuredAsync_RateLimited_ThrowsWithStatusAndRetryAfter()
        {
            var fake = new FakeHttpMessageHandler();
            fake.Queue(_ =>
            {
                var response = new HttpResponseMessage((HttpStatusCode)429) { Content = new StringContent("{\"detail\":\"slow down\"}") };
                response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(7));
                return response;
            });
            using (var authHttp = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(_authPath, authHttp))
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new CodexChatService(auth, chatHttp))
            {
                var ex = await Assert.ThrowsAsync<CodexBackendException>(
                    () => chat.CompleteStructuredAsync("i", TextAndImage(), "page_plan", Schema(), "Low", CancellationToken.None));

                Assert.Equal(429, (int)ex.StatusCode);
                Assert.Equal(TimeSpan.FromSeconds(7), ex.RetryAfter);
                Assert.Contains("slow down", ex.Message);
            }
        }
    }
}
