using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Updates;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class GitHubReleaseClientTests
    {
        private const string SampleJson = @"{
            ""tag_name"": ""v2.1.0"",
            ""html_url"": ""https://github.com/kirklandsig/OutlookAI/releases/tag/v2.1.0"",
            ""published_at"": ""2026-06-02T19:14:00Z"",
            ""body"": ""Release notes here"",
            ""assets"": [
                { ""name"": ""OutlookAI-v2.1.0-RDS-Deploy.zip"",
                  ""browser_download_url"": ""https://github.com/.../OutlookAI-v2.1.0-RDS-Deploy.zip"" },
                { ""name"": ""OutlookAI-v2.1.0-RDS-Deploy.zip.sha256"",
                  ""browser_download_url"": ""https://github.com/.../OutlookAI-v2.1.0-RDS-Deploy.zip.sha256"" }
            ]
        }";

        private static GitHubReleaseClient NewClient(FakeHttpMessageHandler handler)
        {
            var http = new HttpClient(handler);
            return new GitHubReleaseClient(http, "kirklandsig/OutlookAI", "OutlookAI-Updater/test");
        }

        [Fact]
        public async Task GetLatestStableAsync_HappyPath_ReturnsReleaseFound()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.OK, SampleJson);

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            var found = Assert.IsType<ReleaseFound>(result);
            Assert.Equal("v2.1.0", found.Info.Tag);
            Assert.EndsWith("OutlookAI-v2.1.0-RDS-Deploy.zip", found.Info.ZipUrl);
            Assert.EndsWith(".sha256", found.Info.ShaUrl);
            Assert.Equal("OutlookAI-v2.1.0-RDS-Deploy.zip", found.Info.ZipAssetName);
        }

        [Fact]
        public async Task GetLatestStableAsync_404_ReturnsNoReleasesAvailable()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.NotFound, "{\"message\":\"Not Found\"}");

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            Assert.IsType<NoReleasesAvailable>(result);
        }

        [Fact]
        public async Task GetLatestStableAsync_5xx_ReturnsNetworkError()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.InternalServerError, "{}");

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            Assert.IsType<NetworkError>(result);
        }

        [Fact]
        public async Task GetLatestStableAsync_MissingShaAsset_ShaUrlIsNull()
        {
            var noShaJson = @"{
                ""tag_name"": ""v2.1.0"",
                ""html_url"": ""https://github.com/.../releases/tag/v2.1.0"",
                ""published_at"": ""2026-06-02T19:14:00Z"",
                ""body"": """",
                ""assets"": [
                    { ""name"": ""OutlookAI-v2.1.0-RDS-Deploy.zip"",
                      ""browser_download_url"": ""https://github.com/.../zip"" }
                ]
            }";
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.OK, noShaJson);

            var result = await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            var found = Assert.IsType<ReleaseFound>(result);
            Assert.NotNull(found.Info.ZipUrl);
            Assert.Null(found.Info.ShaUrl);
        }

        [Fact]
        public async Task GetLatestStableAsync_SendsUserAgentAndAccept()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueJson(HttpStatusCode.OK, SampleJson);

            await NewClient(handler).GetLatestStableAsync(CancellationToken.None);

            var req = handler.Requests[0];
            Assert.Equal("OutlookAI-Updater", req.Headers.UserAgent.ToString().Split('/')[0]);
            Assert.Contains("application/vnd.github+json",
                req.Headers.Accept.ToString());
            Assert.Equal("https://api.github.com/repos/kirklandsig/OutlookAI/releases/latest",
                req.RequestUri.ToString());
        }
    }
}
