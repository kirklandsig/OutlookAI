using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Models;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Models
{
    public class CodexModelsClientTests
    {
        [Fact]
        public async Task FetchModelsAsync_SendsClientVersionAndTheChatAuthHeaders()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.OK, TestCatalogs.LiveModelsResponse);
            using (var http = new HttpClient(fake))
            {
                var result = await new CodexModelsClient(http)
                    .FetchModelsAsync("tok", "acct-1", "0.155.1", CancellationToken.None);

                Assert.True(result.Succeeded, result.Error);
                Assert.Equal(4, result.Models.Count);
                var req = fake.Requests.Single();
                Assert.Equal(HttpMethod.Get, req.Method);
                Assert.Equal("https://chatgpt.com/backend-api/codex/models?client_version=0.155.1", req.RequestUri.ToString());
                Assert.Equal("Bearer tok", req.Headers.Authorization.ToString());
                Assert.Equal("acct-1", req.Headers.GetValues("ChatGPT-Account-ID").Single());
            }
        }

        [Fact]
        public async Task FetchModelsAsync_OmitsAccountHeader_WhenNoAccountId()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.OK, TestCatalogs.LiveModelsResponse);
            using (var http = new HttpClient(fake))
            {
                await new CodexModelsClient(http).FetchModelsAsync("tok", "", "0.155.1", CancellationToken.None);

                Assert.False(fake.Requests.Single().Headers.Contains("ChatGPT-Account-ID"));
            }
        }

        [Fact]
        public async Task FetchModelsAsync_HttpError_ReportsStatusAndServerMessage()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.Unauthorized, "{\"detail\":\"Could not validate credentials\"}");
            using (var http = new HttpClient(fake))
            {
                var result = await new CodexModelsClient(http).FetchModelsAsync("tok", "a", "0.155.1", CancellationToken.None);

                Assert.False(result.Succeeded);
                Assert.Contains("401", result.Error);
                Assert.Contains("Could not validate credentials", result.Error);
                Assert.Null(result.Models);
            }
        }

        [Theory]
        [InlineData("<html>Bad gateway</html>")]
        [InlineData("{\"data\":[]}")]
        [InlineData("{\"models\":[]}")]
        [InlineData("{\"models\":[{\"slug\":\"hidden-only\",\"visibility\":\"hide\"}]}")]
        public async Task FetchModelsAsync_UnusableBody_IsAFailure(string body)
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueText(HttpStatusCode.OK, body);
            using (var http = new HttpClient(fake))
            {
                var result = await new CodexModelsClient(http).FetchModelsAsync("tok", "a", "0.155.1", CancellationToken.None);

                Assert.False(result.Succeeded);
                Assert.False(string.IsNullOrEmpty(result.Error));
            }
        }

        [Fact]
        public async Task FetchModelsAsync_NoReasoningLevelsParsed_IsAFailure()
        {
            // A renamed/reshaped supported_reasoning_levels must not wipe every
            // effort from the shared cache.
            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.OK,
                "{\"models\":[{\"slug\":\"m\",\"visibility\":\"list\",\"supported_reasoning_levels\":[{\"level\":\"high\"}]}]}");
            using (var http = new HttpClient(fake))
            {
                var result = await new CodexModelsClient(http).FetchModelsAsync("tok", "a", "0.155.1", CancellationToken.None);

                Assert.False(result.Succeeded);
                Assert.Contains("reasoning", result.Error);
            }
        }

        [Fact]
        public async Task FetchModelsAsync_StalledBody_EndsWhenCancelled()
        {
            // On .NET Framework a pending read on the response stream ignores its
            // token; the client has to tear the response down itself.
            var stalling = new StallingStream(System.Text.Encoding.UTF8.GetBytes("{\"models\":["));
            var fake = new FakeHttpMessageHandler();
            fake.QueueRaw(HttpStatusCode.OK, new StreamContent(stalling));
            using (var http = new HttpClient(fake))
            using (var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(300)))
            {
                var fetch = new CodexModelsClient(http).FetchModelsAsync("tok", "a", "0.155.1", cts.Token);
                var finished = await Task.WhenAny(fetch, Task.Delay(TimeSpan.FromSeconds(10)));

                Assert.Same(fetch, finished);
                await Assert.ThrowsAnyAsync<System.OperationCanceledException>(() => fetch);
                Assert.True(stalling.Disposed);
            }
        }

        [Fact]
        public async Task FetchModelsAsync_OversizedBody_IsAFailure()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.OK, TestCatalogs.LiveModelsResponse);
            using (var http = new HttpClient(fake))
            {
                var result = await new CodexModelsClient(http, maxResponseBytes: 100)
                    .FetchModelsAsync("tok", "a", "0.155.1", CancellationToken.None);

                Assert.False(result.Succeeded);
                Assert.Contains("too large", result.Error);
            }
        }

        [Fact]
        public async Task FetchModelsAsync_NetworkFailure_IsAFailure()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueException(new HttpRequestException("No such host is known"));
            using (var http = new HttpClient(fake))
            {
                var result = await new CodexModelsClient(http).FetchModelsAsync("tok", "a", "0.155.1", CancellationToken.None);

                Assert.False(result.Succeeded);
                Assert.Contains("No such host is known", result.Error);
            }
        }

        [Fact]
        public async Task ProbeServerEffortsAsync_SendsAnInvalidEffort_AndParsesTheRejection()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.BadRequest, TestCatalogs.InvalidEffortErrorBody);
            using (var http = new HttpClient(fake))
            {
                var efforts = await new CodexModelsClient(http)
                    .ProbeServerEffortsAsync("tok", "acct-1", "gpt-6-astra", CancellationToken.None);

                Assert.Equal(new[] { "none", "minimal", "low", "medium", "high", "xhigh", "max" }, efforts);
                var req = fake.Requests.Single();
                Assert.Equal(HttpMethod.Post, req.Method);
                Assert.Equal("https://chatgpt.com/backend-api/codex/responses", req.RequestUri.ToString());
                Assert.Equal("Bearer tok", req.Headers.Authorization.ToString());
                var body = JObject.Parse(fake.RequestBodies.Single());
                Assert.Equal("gpt-6-astra", (string)body["model"]);
                Assert.Equal(CodexModelsClient.ProbeEffort, (string)body["reasoning"]["effort"]);
                Assert.False((bool)body["store"]);
            }
        }

        [Fact]
        public async Task ProbeServerEffortsAsync_ReturnsNull_WhenTheServerDoesNotRejectAsExpected()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK, "data: {\"type\":\"response.completed\"}\n\n");
            fake.QueueText(HttpStatusCode.InternalServerError, "oops");
            fake.QueueJson(HttpStatusCode.BadRequest, "{\"detail\":\"The 'x' model is not supported\"}");
            fake.QueueException(new HttpRequestException("reset"));
            using (var http = new HttpClient(fake))
            {
                var client = new CodexModelsClient(http);
                for (int i = 0; i < 4; i++)
                {
                    Assert.Null(await client.ProbeServerEffortsAsync("tok", "a", "m", CancellationToken.None));
                }
            }
        }
    }
}
