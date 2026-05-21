using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Downloads a release ZIP, verifies the SHA256 against a sidecar file,
    /// extracts the archive into the per-tag staging dir, and returns either
    /// a ready-to-launch DownloadSuccess or a typed failure.
    /// </summary>
    public sealed class UpdateDownloader
    {
        private readonly HttpClient _http;

        public UpdateDownloader(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<DownloadResult> DownloadAsync(
            ReleaseInfo info,
            IProgress<int> progress,
            CancellationToken ct)
        {
            if (info == null) throw new ArgumentNullException(nameof(info));
            if (string.IsNullOrWhiteSpace(info.ShaUrl))
            {
                return new DownloadFailed { Detail = "Release is missing the .sha256 sidecar; refusing to download." };
            }

            var stagingDir = Path.Combine(UpdatePaths.BaseUpdatesDir, info.Tag);
            try { if (Directory.Exists(stagingDir)) Directory.Delete(stagingDir, recursive: true); } catch { }
            Directory.CreateDirectory(stagingDir);

            var zipPath = Path.Combine(stagingDir, info.ZipAssetName ?? "release.zip");
            var shaPath = zipPath + ".sha256";

            try
            {
                await DownloadToFileAsync(info.ZipUrl, zipPath, progress, ct).ConfigureAwait(false);
                await DownloadToFileAsync(info.ShaUrl, shaPath, null, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                // Deleting the staging dir here also removes the partially-written zip
                // that DownloadToFileAsync left behind when cancellation tripped the
                // inner loop. Both streams are already disposed (via using) before we
                // get here, so the file handle is released.
                try { Directory.Delete(stagingDir, true); } catch { }
                return new Cancelled();
            }
            catch (Exception ex)
            {
                try { Directory.Delete(stagingDir, true); } catch { }
                return new DownloadFailed { Detail = ex.Message };
            }

            var expected = (File.ReadAllText(shaPath) ?? string.Empty).Trim().ToLowerInvariant();
            var actual = ComputeSha256(zipPath);
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
            {
                try { Directory.Delete(stagingDir, true); } catch { }
                return new HashMismatch { Expected = expected, Actual = actual };
            }

            var extracted = Path.Combine(stagingDir, "extracted");
            try
            {
                if (Directory.Exists(extracted)) Directory.Delete(extracted, true);
                ZipFile.ExtractToDirectory(zipPath, extracted);
            }
            catch (Exception ex)
            {
                try { Directory.Delete(stagingDir, true); } catch { }
                return new DownloadFailed { Detail = "Extract failed: " + ex.Message };
            }

            var installer = Path.Combine(extracted, "Install-OutlookAI.ps1");
            if (!File.Exists(installer))
            {
                return new MissingInstallerScript();
            }

            return new DownloadSuccess
            {
                StagingDir = stagingDir,
                ExtractedDir = extracted,
                InstallerScriptPath = installer,
                ExpectedSha256 = expected,
            };
        }

        private async Task DownloadToFileAsync(string url, string destPath, IProgress<int> progress, CancellationToken ct)
        {
            using (var resp = await _http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
            {
                if (!resp.IsSuccessStatusCode)
                {
                    throw new HttpRequestException("HTTP " + (int)resp.StatusCode + " from " + url);
                }

                var total = resp.Content.Headers.ContentLength ?? -1L;
                using (var src = await resp.Content.ReadAsStreamAsync().ConfigureAwait(false))
                using (var dst = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None, 64 * 1024, useAsync: true))
                {
                    var buffer = new byte[64 * 1024];
                    long written = 0;
                    int n;
                    while ((n = await src.ReadAsync(buffer, 0, buffer.Length, ct).ConfigureAwait(false)) > 0)
                    {
                        await dst.WriteAsync(buffer, 0, n, ct).ConfigureAwait(false);
                        written += n;
                        if (progress != null && total > 0)
                        {
                            progress.Report((int)(written * 100 / total));
                        }
                    }
                }
            }
        }

        private static string ComputeSha256(string path)
        {
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(path))
            {
                var hash = sha.ComputeHash(stream);
                var sb = new StringBuilder(hash.Length * 2);
                foreach (var b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }
    }
}
