using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    /// <summary>
    /// Authentication mode reported by <see cref="CodexAuthService"/>.
    /// </summary>
    public enum AuthState
    {
        Unauthenticated,
        Authenticated,
        Error
    }

    /// <summary>
    /// Snapshot of the current auth state plus the user-facing email/error,
    /// surfaced to the Settings UI through <see cref="CodexAuthService.GetStatus"/>
    /// and <see cref="CodexAuthService.StatusChanged"/>.
    /// </summary>
    public sealed class AuthStatus
    {
        public AuthState State { get; }
        public string Email { get; }
        public string AccountId { get; }
        public string Message { get; }

        public AuthStatus(AuthState state, string email, string accountId, string message)
        {
            State = state;
            Email = email ?? "";
            AccountId = accountId ?? "";
            Message = message ?? "";
        }

        public static AuthStatus Unauthenticated(string message = "Not signed in")
            => new AuthStatus(AuthState.Unauthenticated, "", "", message);

        public static AuthStatus FromTokens(StoredTokens tokens)
            => new AuthStatus(AuthState.Authenticated, tokens.Email, tokens.AccountId, "");

        public static AuthStatus Error(string message)
            => new AuthStatus(AuthState.Error, "", "", message);
    }

    /// <summary>
    /// Persisted token bundle. Mirrors the shape of <c>~/.codex/auth.json</c>
    /// produced by Codex CLI; OutlookAI writes it to <c>C:\ProgramData\OutlookAI\auth.json</c>.
    /// </summary>
    public sealed class StoredTokens
    {
        public string AccessToken { get; set; } = "";
        public string IdToken { get; set; } = "";
        public string RefreshToken { get; set; } = "";
        public string AccountId { get; set; } = "";
        public string Email { get; set; } = "";
        public DateTimeOffset? AccessTokenExpiresAt { get; set; }
        public DateTimeOffset SavedAt { get; set; } = DateTimeOffset.UtcNow;
    }

    /// <summary>
    /// Owns the embedded ChatGPT/Codex OAuth flow, on-disk token persistence,
    /// and refresh. Both the chat and voice services consume the OAuth
    /// <c>access_token</c> returned by <see cref="GetAccessTokenAsync"/>.
    ///
    /// Phase 1 reuses the public Codex CLI <c>client_id</c>; the OAuth grant
    /// covers ChatGPT consumer-subscription routing for both
    /// <c>chatgpt.com/backend-api/codex/responses</c> and
    /// <c>wss://api.openai.com/v1/realtime</c>.
    /// </summary>
    public sealed class CodexAuthService : IDisposable
    {
        public const string ClientId = "app_EMoamEEZ73f0CkXaXp7hrann";
        public const string AuthorizeEndpoint = "https://auth.openai.com/oauth/authorize";
        public const string TokenEndpoint = "https://auth.openai.com/oauth/token";
        public const string Scopes = "openid profile email offline_access api.connectors.read api.connectors.invoke";

        // Refresh slightly before the server-side expiry to absorb clock skew.
        private static readonly TimeSpan RefreshSkew = TimeSpan.FromSeconds(60);
        private static readonly int[] CallbackPorts = { 1455, 1457 };

        private readonly string _authPath;
        private readonly string _lockPath;
        private readonly HttpClient _http;
        private readonly bool _ownsHttp;
        private readonly SemaphoreSlim _gate = new SemaphoreSlim(1, 1);

        private StoredTokens _cached;
        private bool _disposed;

        public event EventHandler<AuthStatus> StatusChanged;

        public CodexAuthService(string authPath)
            : this(authPath, BuildDefaultHttpClient(), ownsHttp: true)
        {
        }

        // Test seam.
        public CodexAuthService(string authPath, HttpClient httpClient, bool ownsHttp = false)
        {
            if (string.IsNullOrWhiteSpace(authPath))
            {
                throw new ArgumentException("Auth file path must be provided.", nameof(authPath));
            }
            _authPath = authPath;
            _lockPath = authPath + ".refresh.lock";
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _ownsHttp = ownsHttp;
        }

        private static HttpClient BuildDefaultHttpClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            return new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(60)
            };
        }

        /// <summary>
        /// Returns the current ChatGPT OAuth access token. Refreshes
        /// transparently when expired. Throws <see cref="InvalidOperationException"/>
        /// when no tokens are persisted (caller should prompt for sign-in).
        /// </summary>
        public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                EnsureLoaded();
                if (_cached == null)
                {
                    throw new InvalidOperationException("OutlookAI is not signed in. Open Settings to sign in.");
                }

                if (NeedsRefresh(_cached))
                {
                    var refreshed = await RefreshAsync(_cached.RefreshToken, cancellationToken).ConfigureAwait(false);
                    PersistTokens(refreshed);
                }

                return _cached.AccessToken;
            }
            finally
            {
                _gate.Release();
            }
        }

        public AuthStatus GetStatus()
        {
            try
            {
                EnsureLoaded();
                return _cached == null
                    ? AuthStatus.Unauthenticated()
                    : AuthStatus.FromTokens(_cached);
            }
            catch (Exception ex)
            {
                return AuthStatus.Error(ex.Message);
            }
        }

        /// <summary>
        /// Runs the embedded browser OAuth + authorization-code exchange and
        /// persists the result. Process-wide single-flight via <see cref="_gate"/>.
        /// </summary>
        public async Task SignInAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                var (port, listener) = StartCallbackListener();
                try
                {
                    var redirectUri = "http://localhost:" + port + "/auth/callback";
                    var state = RandomBase64Url(32);
                    var verifier = RandomBase64Url(64);
                    var challenge = ComputeCodeChallenge(verifier);
                    var url = BuildAuthorizeUrl(redirectUri, state, challenge);

                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });

                    var callback = await WaitForCallbackAsync(listener, TimeSpan.FromMinutes(5), cancellationToken).ConfigureAwait(false);
                    if (!StringComparer.Ordinal.Equals(callback.State, state))
                    {
                        throw new InvalidOperationException("OAuth state mismatch.");
                    }

                    var tokens = await ExchangeCodeAsync(callback.Code, redirectUri, verifier, cancellationToken).ConfigureAwait(false);
                    PersistTokens(tokens);
                }
                finally
                {
                    try { listener.Stop(); } catch { /* ignore */ }
                    try { listener.Close(); } catch { /* ignore */ }
                }
            }
            catch (Exception ex)
            {
                RaiseStatusChanged(AuthStatus.Error(ex.Message));
                throw;
            }
            finally
            {
                _gate.Release();
            }
        }

        public async Task SignOutAsync()
        {
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                _cached = null;
                try
                {
                    if (File.Exists(_authPath))
                    {
                        File.Delete(_authPath);
                    }
                }
                catch
                {
                    // Best-effort delete.
                }
                try
                {
                    if (File.Exists(_lockPath))
                    {
                        File.Delete(_lockPath);
                    }
                }
                catch
                {
                    // Best-effort delete.
                }
            }
            finally
            {
                _gate.Release();
            }

            RaiseStatusChanged(AuthStatus.Unauthenticated("Signed out"));
        }

        // -------------------------------------------------------------------
        // OAuth helpers
        // -------------------------------------------------------------------

        private static (int Port, HttpListener Listener) StartCallbackListener()
        {
            Exception last = null;
            foreach (var port in CallbackPorts)
            {
                try
                {
                    var listener = new HttpListener();
                    listener.Prefixes.Add("http://localhost:" + port + "/auth/");
                    listener.Start();
                    return (port, listener);
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }
            throw new InvalidOperationException(
                "Could not bind any OAuth callback port (" + string.Join(", ", CallbackPorts) + "): " + (last == null ? "" : last.Message),
                last);
        }

        private static async Task<CallbackResult> WaitForCallbackAsync(HttpListener listener, TimeSpan timeout, CancellationToken cancellationToken)
        {
            var contextTask = listener.GetContextAsync();
            using (var delayCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                var delayTask = Task.Delay(timeout, delayCts.Token);
                var winner = await Task.WhenAny(contextTask, delayTask).ConfigureAwait(false);
                if (winner != contextTask)
                {
                    throw new TimeoutException("Timed out waiting for OAuth callback.");
                }
                delayCts.Cancel();
            }

            var context = await contextTask.ConfigureAwait(false);
            var query = context.Request.QueryString;
            var result = new CallbackResult
            {
                Code = query["code"] ?? "",
                State = query["state"] ?? ""
            };

            var body = Encoding.UTF8.GetBytes(
                "<html><head><title>OutlookAI</title></head>" +
                "<body style=\"font-family:Segoe UI,Arial;padding:32px;\">" +
                "<h2>OutlookAI sign-in complete.</h2>" +
                "<p>You can close this tab and return to Outlook.</p>" +
                "</body></html>");
            try
            {
                context.Response.ContentType = "text/html";
                context.Response.ContentLength64 = body.Length;
                context.Response.OutputStream.Write(body, 0, body.Length);
            }
            finally
            {
                try { context.Response.Close(); } catch { /* ignore */ }
            }

            if (string.IsNullOrEmpty(result.Code))
            {
                throw new InvalidOperationException("OAuth callback did not include an authorization code.");
            }
            return result;
        }

        private static string BuildAuthorizeUrl(string redirectUri, string state, string codeChallenge)
        {
            var values = new Dictionary<string, string>
            {
                { "client_id", ClientId },
                { "response_type", "code" },
                { "redirect_uri", redirectUri },
                { "scope", Scopes },
                { "state", state },
                { "code_challenge", codeChallenge },
                { "code_challenge_method", "S256" }
            };
            var query = string.Join("&", values.Select(kv => Uri.EscapeDataString(kv.Key) + "=" + Uri.EscapeDataString(kv.Value)));
            return AuthorizeEndpoint + "?" + query;
        }

        private async Task<StoredTokens> ExchangeCodeAsync(string code, string redirectUri, string codeVerifier, CancellationToken cancellationToken)
        {
            var body = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                { "grant_type", "authorization_code" },
                { "client_id", ClientId },
                { "code", code },
                { "redirect_uri", redirectUri },
                { "code_verifier", codeVerifier }
            });

            using (var response = await _http.PostAsync(TokenEndpoint, body, cancellationToken).ConfigureAwait(false))
            {
                var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException("OAuth code exchange failed: " + (int)response.StatusCode + " " + text);
                }
                return ParseTokenResponse(text);
            }
        }

        private async Task<StoredTokens> RefreshAsync(string refreshToken, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(refreshToken))
            {
                throw new InvalidOperationException("Cannot refresh: refresh_token is missing. Sign in again.");
            }

            var body = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                { "grant_type", "refresh_token" },
                { "client_id", ClientId },
                { "refresh_token", refreshToken }
            });

            using (var response = await _http.PostAsync(TokenEndpoint, body, cancellationToken).ConfigureAwait(false))
            {
                var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException("OAuth refresh failed: " + (int)response.StatusCode + " " + text);
                }
                var fresh = ParseTokenResponse(text);
                // OpenAI sometimes omits refresh_token from refresh responses; keep
                // the previous one when that happens so we can keep refreshing.
                if (string.IsNullOrEmpty(fresh.RefreshToken))
                {
                    fresh.RefreshToken = refreshToken;
                }
                return fresh;
            }
        }

        private static StoredTokens ParseTokenResponse(string json)
        {
            var obj = JObject.Parse(json);
            var accessToken = (string)obj["access_token"];
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new InvalidOperationException("OAuth response did not include access_token.");
            }
            var idToken = (string)obj["id_token"] ?? "";
            var refreshToken = (string)obj["refresh_token"] ?? "";
            var expiresIn = (int?)obj["expires_in"];
            var (email, accountId) = ExtractIdTokenClaims(idToken);

            return new StoredTokens
            {
                AccessToken = accessToken,
                IdToken = idToken,
                RefreshToken = refreshToken,
                Email = email,
                AccountId = accountId,
                AccessTokenExpiresAt = expiresIn.HasValue
                    ? DateTimeOffset.UtcNow.AddSeconds(expiresIn.Value)
                    : (DateTimeOffset?)null,
                SavedAt = DateTimeOffset.UtcNow
            };
        }

        private static (string Email, string AccountId) ExtractIdTokenClaims(string idToken)
        {
            if (string.IsNullOrEmpty(idToken))
            {
                return ("", "");
            }
            var parts = idToken.Split('.');
            if (parts.Length < 2)
            {
                return ("", "");
            }
            var payload = parts[1].Replace('-', '+').Replace('_', '/');
            switch (payload.Length % 4)
            {
                case 2: payload += "=="; break;
                case 3: payload += "="; break;
            }
            try
            {
                var json = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                var obj = JObject.Parse(json);
                var email = (string)obj["email"] ?? "";
                var auth = obj["https://api.openai.com/auth"] as JObject;
                var accountId = auth != null ? (string)auth["chatgpt_account_id"] ?? "" : "";
                return (email, accountId);
            }
            catch
            {
                return ("", "");
            }
        }

        private static string RandomBase64Url(int byteCount)
        {
            var bytes = new byte[byteCount];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }
            return Base64Url(bytes);
        }

        private static string ComputeCodeChallenge(string verifier)
        {
            using (var sha = SHA256.Create())
            {
                var hash = sha.ComputeHash(Encoding.ASCII.GetBytes(verifier));
                return Base64Url(hash);
            }
        }

        private static string Base64Url(byte[] bytes)
        {
            return Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }

        private bool NeedsRefresh(StoredTokens tokens)
        {
            if (tokens == null || string.IsNullOrEmpty(tokens.AccessToken))
            {
                return true;
            }
            if (!tokens.AccessTokenExpiresAt.HasValue)
            {
                // No expiry hint: try to read the JWT; if we can't, accept the
                // token until OpenAI tells us otherwise.
                var jwtExpiry = TryReadJwtExpiry(tokens.AccessToken);
                if (!jwtExpiry.HasValue)
                {
                    return false;
                }
                tokens.AccessTokenExpiresAt = jwtExpiry;
            }
            return DateTimeOffset.UtcNow + RefreshSkew >= tokens.AccessTokenExpiresAt.Value;
        }

        private static DateTimeOffset? TryReadJwtExpiry(string jwt)
        {
            var parts = jwt.Split('.');
            if (parts.Length < 2)
            {
                return null;
            }
            var payload = parts[1].Replace('-', '+').Replace('_', '/');
            switch (payload.Length % 4)
            {
                case 2: payload += "=="; break;
                case 3: payload += "="; break;
            }
            try
            {
                var json = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                var obj = JObject.Parse(json);
                var exp = (long?)obj["exp"];
                return exp.HasValue ? DateTimeOffset.FromUnixTimeSeconds(exp.Value) : (DateTimeOffset?)null;
            }
            catch
            {
                return null;
            }
        }

        // -------------------------------------------------------------------
        // Persistence
        // -------------------------------------------------------------------

        private void EnsureLoaded()
        {
            if (_cached != null)
            {
                return;
            }
            if (!File.Exists(_authPath))
            {
                return;
            }
            try
            {
                var json = File.ReadAllText(_authPath, Encoding.UTF8);
                _cached = ParseStoredJson(json);
            }
            catch
            {
                _cached = null;
            }
        }

        private static StoredTokens ParseStoredJson(string json)
        {
            var obj = JObject.Parse(json);
            var tokens = obj["tokens"] as JObject ?? obj;
            return new StoredTokens
            {
                AccessToken = (string)tokens["access_token"] ?? "",
                IdToken = (string)tokens["id_token"] ?? "",
                RefreshToken = (string)tokens["refresh_token"] ?? "",
                AccountId = (string)tokens["account_id"] ?? "",
                Email = (string)tokens["email"] ?? "",
                AccessTokenExpiresAt = TryParseDate((string)tokens["access_token_expires_at"]),
                SavedAt = TryParseDate((string)obj["last_refresh"]) ?? DateTimeOffset.UtcNow
            };
        }

        private static DateTimeOffset? TryParseDate(string raw)
        {
            DateTimeOffset parsed;
            return DateTimeOffset.TryParse(raw, out parsed) ? parsed : (DateTimeOffset?)null;
        }

        private void PersistTokens(StoredTokens tokens)
        {
            var dir = Path.GetDirectoryName(_authPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var payload = new JObject(
                new JProperty("tokens", new JObject(
                    new JProperty("access_token", tokens.AccessToken),
                    new JProperty("id_token", tokens.IdToken),
                    new JProperty("refresh_token", tokens.RefreshToken),
                    new JProperty("account_id", tokens.AccountId),
                    new JProperty("email", tokens.Email),
                    new JProperty(
                        "access_token_expires_at",
                        tokens.AccessTokenExpiresAt.HasValue
                            ? tokens.AccessTokenExpiresAt.Value.ToString("o")
                            : null))),
                new JProperty("last_refresh", DateTimeOffset.UtcNow.ToString("o")));

            var temp = _authPath + ".tmp";
            File.WriteAllText(temp, payload.ToString(Newtonsoft.Json.Formatting.Indented), new UTF8Encoding(false));
            try
            {
                if (File.Exists(_authPath))
                {
                    File.Replace(temp, _authPath, null);
                }
                else
                {
                    File.Move(temp, _authPath);
                }
            }
            catch
            {
                if (File.Exists(temp))
                {
                    try { File.Delete(temp); } catch { /* ignore */ }
                }
                throw;
            }

            _cached = tokens;
            RaiseStatusChanged(AuthStatus.FromTokens(tokens));
        }

        private void RaiseStatusChanged(AuthStatus status)
        {
            var handler = StatusChanged;
            if (handler == null)
            {
                return;
            }
            try
            {
                handler(this, status);
            }
            catch
            {
                // Subscriber bugs must not break auth.
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_ownsHttp)
            {
                _http.Dispose();
            }
            _gate.Dispose();
        }

        private sealed class CallbackResult
        {
            public string Code { get; set; }
            public string State { get; set; }
        }
    }
}
