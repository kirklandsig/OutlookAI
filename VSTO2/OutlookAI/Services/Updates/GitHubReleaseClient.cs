using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Thin async wrapper over HttpClient. Calls api.github.com and parses
    /// the latest-release JSON shape into a ReleaseInfo. Honors the system
    /// proxy via the default HttpClient configuration on .NET Framework.
    /// </summary>
    public sealed class GitHubReleaseClient
    {
        private readonly HttpClient _http;
        private readonly string _repo;
        private readonly string _userAgent;

        public GitHubReleaseClient(HttpClient http, string repo, string userAgent)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
            _repo = repo ?? throw new ArgumentNullException(nameof(repo));
            _userAgent = string.IsNullOrWhiteSpace(userAgent) ? "OutlookAI-Updater/dev" : userAgent;
        }

        public async Task<ReleaseLookupResult> GetLatestStableAsync(CancellationToken ct)
        {
            var url = "https://api.github.com/repos/" + _repo + "/releases/latest";
            var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.UserAgent.ParseAdd(_userAgent);
            req.Headers.Accept.ParseAdd("application/vnd.github+json");

            HttpResponseMessage resp;
            try
            {
                resp = await _http.SendAsync(req, ct).ConfigureAwait(false);
            }
            catch (HttpRequestException ex)
            {
                return new NetworkError { Detail = ex.Message };
            }
            catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
            {
                return new NetworkError { Detail = "Request timed out: " + ex.Message };
            }

            if (resp.StatusCode == HttpStatusCode.NotFound)
            {
                return new NoReleasesAvailable();
            }

            if (resp.StatusCode == HttpStatusCode.Forbidden &&
                resp.Headers.TryGetValues("X-RateLimit-Remaining", out var rem) &&
                rem.FirstOrDefault() == "0")
            {
                var reset = DateTimeOffset.UtcNow.AddHours(1);
                if (resp.Headers.TryGetValues("X-RateLimit-Reset", out var resetVals) &&
                    long.TryParse(resetVals.FirstOrDefault(), out var unixReset))
                {
                    reset = DateTimeOffset.FromUnixTimeSeconds(unixReset);
                }
                return new RateLimited { ResetAt = reset };
            }

            if (!resp.IsSuccessStatusCode)
            {
                return new NetworkError { Detail = "HTTP " + (int)resp.StatusCode };
            }

            var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
            try
            {
                return new ReleaseFound { Info = ParseRelease(body) };
            }
            catch (Exception ex)
            {
                return new NetworkError { Detail = "Malformed release JSON: " + ex.Message };
            }
        }

        internal static ReleaseInfo ParseRelease(string json)
        {
            var o = JObject.Parse(json);
            var info = new ReleaseInfo
            {
                Tag = (string)o["tag_name"],
                ReleasePageUrl = (string)o["html_url"],
                Body = (string)o["body"] ?? string.Empty,
                PublishedAt = ParseDate((string)o["published_at"]),
            };

            var assets = o["assets"] as JArray;
            if (assets != null)
            {
                foreach (var a in assets.OfType<JObject>())
                {
                    var name = (string)a["name"] ?? string.Empty;
                    var url = (string)a["browser_download_url"];
                    if (string.IsNullOrEmpty(url)) continue;

                    if (name.EndsWith(".zip.sha256", StringComparison.OrdinalIgnoreCase))
                    {
                        info.ShaUrl = url;
                    }
                    else if (name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        info.ZipAssetName = name;
                        info.ZipUrl = url;
                    }
                }
            }
            return info;
        }

        private static DateTimeOffset ParseDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return DateTimeOffset.MinValue;
            if (DateTimeOffset.TryParse(s, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal,
                out var dt)) return dt;
            return DateTimeOffset.MinValue;
        }
    }
}
