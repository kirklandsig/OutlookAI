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
    public class CodexAuthServiceTests : IDisposable
    {
        private readonly string _tmpDir;
        private readonly string _authPath;

        public CodexAuthServiceTests()
        {
            _tmpDir = Path.Combine(Path.GetTempPath(), "outlookai-auth", Path.GetRandomFileName());
            Directory.CreateDirectory(_tmpDir);
            _authPath = Path.Combine(_tmpDir, "auth.json");
        }

        public void Dispose()
        {
            try { Directory.Delete(_tmpDir, recursive: true); } catch { }
        }

        [Fact]
        public void GetStatus_ReturnsUnauthenticated_WhenNoAuthFile()
        {
            using (var http = new HttpClient(new FakeHttpMessageHandler()))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                Assert.Equal(AuthState.Unauthenticated, svc.GetStatus().State);
            }
        }

        [Fact]
        public async Task GetAccessTokenAsync_Throws_WhenNotSignedIn()
        {
            using (var http = new HttpClient(new FakeHttpMessageHandler()))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                await Assert.ThrowsAsync<InvalidOperationException>(
                    () => svc.GetAccessTokenAsync(CancellationToken.None));
            }
        }

        [Fact]
        public async Task GetAccessTokenAsync_ReturnsCachedToken_WhenFresh()
        {
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk-fresh\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");

            var fake = new FakeHttpMessageHandler();
            using (var http = new HttpClient(fake))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                var token = await svc.GetAccessTokenAsync(CancellationToken.None);

                Assert.Equal("sk-fresh", token);
                Assert.Empty(fake.Requests);
            }
        }

        [Fact]
        public async Task GetAccessTokenAsync_RefreshesExpiredToken()
        {
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk-old\",\"id_token\":\"\","
                + "\"refresh_token\":\"r1\",\"access_token_expires_at\":\""
                + DateTimeOffset.UtcNow.AddHours(-1).ToString("o")
                + "\"},\"last_refresh\":\"" + DateTimeOffset.UtcNow.ToString("o") + "\"}");

            var fake = new FakeHttpMessageHandler();
            fake.QueueJson(HttpStatusCode.OK,
                "{\"access_token\":\"sk-new\",\"refresh_token\":\"r2\",\"expires_in\":3600}");

            using (var http = new HttpClient(fake))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                var token = await svc.GetAccessTokenAsync(CancellationToken.None);

                Assert.Equal("sk-new", token);
                Assert.Single(fake.Requests);
                Assert.Equal("https://auth.openai.com/oauth/token",
                    fake.Requests[0].RequestUri.ToString());
                Assert.Contains("grant_type=refresh_token", fake.RequestBodies[0]);
                Assert.Contains("refresh_token=r1", fake.RequestBodies[0]);
            }
        }

        [Fact]
        public async Task SignOutAsync_DeletesAuthFile()
        {
            File.WriteAllText(_authPath,
                "{\"tokens\":{\"access_token\":\"sk\",\"refresh_token\":\"r\","
                + "\"id_token\":\"\"}}");
            using (var http = new HttpClient(new FakeHttpMessageHandler()))
            using (var svc = new CodexAuthService(_authPath, http))
            {
                await svc.SignOutAsync();

                Assert.False(File.Exists(_authPath));
                Assert.Equal(AuthState.Unauthenticated, svc.GetStatus().State);
            }
        }
    }
}
