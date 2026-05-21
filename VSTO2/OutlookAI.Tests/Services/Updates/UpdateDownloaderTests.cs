using System;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Updates;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    [Collection("UpdatePaths")]
    public class UpdateDownloaderTests : IDisposable
    {
        private readonly string _tempBase;
        private readonly string _originalBaseDir;

        public UpdateDownloaderTests()
        {
            _tempBase = Path.Combine(Path.GetTempPath(), "updater-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(_tempBase);
            _originalBaseDir = UpdatePaths.BaseUpdatesDir;
            UpdatePaths.BaseUpdatesDir = Path.Combine(_tempBase, "Updates");
        }

        public void Dispose()
        {
            UpdatePaths.BaseUpdatesDir = _originalBaseDir;
            try { Directory.Delete(_tempBase, true); } catch { }
        }

        private static byte[] BuildZip(params (string name, string content)[] entries)
        {
            using (var ms = new MemoryStream())
            {
                using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
                {
                    foreach (var (name, content) in entries)
                    {
                        var e = archive.CreateEntry(name);
                        using (var w = new StreamWriter(e.Open(), Encoding.UTF8)) w.Write(content);
                    }
                }
                return ms.ToArray();
            }
        }

        private static string Sha256(byte[] data)
        {
            using (var sha = SHA256.Create())
            {
                var hash = sha.ComputeHash(data);
                var sb = new StringBuilder();
                foreach (var b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        private static ReleaseInfo Info(string tag = "v2.1.0") => new ReleaseInfo
        {
            Tag = tag,
            ZipAssetName = "OutlookAI-" + tag + "-RDS-Deploy.zip",
            ZipUrl = "https://github.test/zip",
            ShaUrl = "https://github.test/sha",
        };

        [Fact]
        public async Task DownloadAsync_HashMatches_ReturnsSuccessWithInstallerPath()
        {
            var zip = BuildZip(
                ("Install-OutlookAI.ps1", "Write-Host hi"),
                ("version.json", "{\"tag\":\"v2.1.0\"}"));
            var sha = Sha256(zip);

            var handler = new FakeHttpMessageHandler();
            handler.QueueRaw(HttpStatusCode.OK, new ByteArrayContent(zip));
            handler.QueueText(HttpStatusCode.OK, sha);

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            var ok = Assert.IsType<DownloadSuccess>(result);
            Assert.True(File.Exists(ok.InstallerScriptPath));
            Assert.EndsWith(@"\v2.1.0\extracted\Install-OutlookAI.ps1", ok.InstallerScriptPath);
            Assert.Equal(sha, ok.ExpectedSha256);
        }

        [Fact]
        public async Task DownloadAsync_HashMismatch_DeletesStagingAndReturnsHashMismatch()
        {
            var zip = BuildZip(("Install-OutlookAI.ps1", "Write-Host hi"));
            var wrong = new string('0', 64);

            var handler = new FakeHttpMessageHandler();
            handler.QueueRaw(HttpStatusCode.OK, new ByteArrayContent(zip));
            handler.QueueText(HttpStatusCode.OK, wrong);

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            Assert.IsType<HashMismatch>(result);
            var stagingDir = Path.Combine(UpdatePaths.BaseUpdatesDir, "v2.1.0");
            Assert.False(Directory.Exists(stagingDir));
        }

        [Fact]
        public async Task DownloadAsync_ZipMissingInstaller_ReturnsMissingInstallerScript()
        {
            var zip = BuildZip(("readme.txt", "no installer here"));
            var sha = Sha256(zip);

            var handler = new FakeHttpMessageHandler();
            handler.QueueRaw(HttpStatusCode.OK, new ByteArrayContent(zip));
            handler.QueueText(HttpStatusCode.OK, sha);

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            Assert.IsType<MissingInstallerScript>(result);
        }

        [Fact]
        public async Task DownloadAsync_HttpError_ReturnsDownloadFailed()
        {
            var handler = new FakeHttpMessageHandler();
            handler.QueueText(HttpStatusCode.InternalServerError, "boom");

            var downloader = new UpdateDownloader(new HttpClient(handler));
            var result = await downloader.DownloadAsync(Info(), null, CancellationToken.None);

            Assert.IsType<DownloadFailed>(result);
        }

        [Fact]
        public async Task DownloadAsync_NullShaUrl_ReturnsDownloadFailed()
        {
            var handler = new FakeHttpMessageHandler();
            var downloader = new UpdateDownloader(new HttpClient(handler));
            var info = Info();
            info.ShaUrl = null;

            var result = await downloader.DownloadAsync(info, null, CancellationToken.None);

            var failed = Assert.IsType<DownloadFailed>(result);
            Assert.Contains("sha256", failed.Detail, StringComparison.OrdinalIgnoreCase);
        }
    }
}
