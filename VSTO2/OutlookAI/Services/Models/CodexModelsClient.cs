using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.Models
{
    public sealed class ModelsFetchResult
    {
        private ModelsFetchResult(IReadOnlyList<ModelCatalogEntry> models, string error)
        {
            Models = models;
            Error = error;
        }

        public bool Succeeded => Models != null;

        /// <summary>Parsed catalog entries; null on failure.</summary>
        public IReadOnlyList<ModelCatalogEntry> Models { get; }

        /// <summary>User-facing failure reason; null on success.</summary>
        public string Error { get; }

        internal static ModelsFetchResult Ok(IReadOnlyList<ModelCatalogEntry> models) => new ModelsFetchResult(models, null);
        internal static ModelsFetchResult Fail(string error) => new ModelsFetchResult(null, error);
    }

    /// <summary>
    /// The two Codex backend calls a catalog refresh makes, with the same auth
    /// headers as <see cref="CodexChatService"/>: the model catalog
    /// (<c>GET /models</c>, what Codex CLI caches in ~/.codex/models_cache.json)
    /// and a one-shot reasoning-effort probe against <c>/responses</c>.
    /// </summary>
    public sealed class CodexModelsClient
    {
        public const string ModelsEndpoint = "https://chatgpt.com/backend-api/codex/models";

        /// <summary>
        /// Deliberately invalid effort. /responses rejects it with a 400 that lists
        /// the valid set, before any inference runs.
        /// </summary>
        public const string ProbeEffort = "outlookai_probe";

        public const int DefaultMaxResponseBytes = 8 * 1024 * 1024;
        private const int MaxErrorBodyBytes = 64 * 1024;

        private readonly HttpClient _http;
        private readonly int _maxResponseBytes;

        public CodexModelsClient(HttpClient http)
            : this(http, DefaultMaxResponseBytes)
        {
        }

        internal CodexModelsClient(HttpClient http, int maxResponseBytes)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
            _maxResponseBytes = maxResponseBytes;
        }

        public async Task<ModelsFetchResult> FetchModelsAsync(
            string accessToken, string accountId, string clientVersion, CancellationToken ct)
        {
            var url = ModelsEndpoint + "?client_version=" + Uri.EscapeDataString(clientVersion ?? "");
            string body;
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                {
                    AddAuthHeaders(request, accessToken, accountId);
                    request.Headers.Accept.ParseAdd("application/json");
                    using (var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            var errorBody = await ReadBodyAsync(response, MaxErrorBodyBytes, ct).ConfigureAwait(false);
                            return ModelsFetchResult.Fail(
                                "ChatGPT returned HTTP " + (int)response.StatusCode + DescribeServerError(errorBody) + ".");
                        }
                        body = await ReadBodyAsync(response, _maxResponseBytes, ct).ConfigureAwait(false);
                    }
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                return ModelsFetchResult.Fail("The request to ChatGPT timed out.");
            }
            catch (Exception ex) when (ex is HttpRequestException || ex is IOException)
            {
                return ModelsFetchResult.Fail("Could not reach ChatGPT: " + ex.GetBaseException().Message);
            }

            if (body == null)
            {
                return ModelsFetchResult.Fail("The model list from ChatGPT was too large to accept.");
            }

            List<ModelCatalogEntry> models;
            try
            {
                models = ModelCatalogJson.ParseModelsResponse(body);
            }
            catch (Exception ex) when (ex is JsonException || ex is FormatException)
            {
                return ModelsFetchResult.Fail("ChatGPT sent a model list OutlookAI could not read (" + ex.Message + ").");
            }
            if (!models.Any(m => m.Listed))
            {
                return ModelsFetchResult.Fail("ChatGPT returned no selectable models for this account.");
            }
            // Codex models all reason; none with a readable level means the
            // format changed, and saving it would wipe every effort setting.
            if (!models.Any(m => m.Listed && m.Efforts.Count > 0))
            {
                return ModelsFetchResult.Fail("ChatGPT's model list had no reasoning levels OutlookAI could read.");
            }
            return ModelsFetchResult.Ok(models.AsReadOnly());
        }

        /// <summary>
        /// The server's global reasoning.effort set, read from its rejection of
        /// <see cref="ProbeEffort"/>; null when the reply isn't that 400.
        /// </summary>
        public async Task<IReadOnlyList<string>> ProbeServerEffortsAsync(
            string accessToken, string accountId, string model, CancellationToken ct)
        {
            var payload = new JObject(
                new JProperty("model", model),
                new JProperty("instructions", "Reply with OK."),
                new JProperty("input", new JArray(new JObject(
                    new JProperty("type", "message"),
                    new JProperty("role", "user"),
                    new JProperty("content", "ping")))),
                new JProperty("tools", new JArray()),
                new JProperty("tool_choice", "auto"),
                new JProperty("parallel_tool_calls", false),
                new JProperty("reasoning", new JObject(new JProperty("effort", ProbeEffort))),
                new JProperty("store", false),
                new JProperty("stream", true),
                new JProperty("include", new JArray()));
            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, CodexChatService.ResponsesEndpoint))
                {
                    AddAuthHeaders(request, accessToken, accountId);
                    request.Headers.Accept.ParseAdd("text/event-stream");
                    request.Content = new StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json");
                    using (var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                    {
                        // Anything but the expected 400 tells us nothing; never read a success stream.
                        if ((int)response.StatusCode != 400)
                        {
                            TraceLog.Write("Effort probe: unexpected HTTP " + (int)response.StatusCode, "Models");
                            return null;
                        }
                        var body = await ReadBodyAsync(response, MaxErrorBodyBytes, ct).ConfigureAwait(false);
                        return ModelCatalogJson.ParseSupportedEffortsFromError(body);
                    }
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                TraceLog.Write("Effort probe failed: " + ex.GetBaseException().Message, "Models");
                return null;
            }
        }

        private static void AddAuthHeaders(HttpRequestMessage request, string accessToken, string accountId)
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            if (!string.IsNullOrEmpty(accountId))
            {
                request.Headers.TryAddWithoutValidation("ChatGPT-Account-ID", accountId);
            }
        }

        // Null when the body exceeds maxBytes. On .NET Framework a pending read on
        // the response stream ignores its token, so a body that stalls midway
        // would hang forever; cancellation disposes the response to end it.
        private static async Task<string> ReadBodyAsync(HttpResponseMessage response, int maxBytes, CancellationToken ct)
        {
            var content = response.Content;
            if (content == null) return "";
            var declared = content.Headers.ContentLength;
            if (declared.HasValue && declared.Value > maxBytes) return null;
            using (ct.Register(response.Dispose))
            {
                try
                {
                    using (var stream = await content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var buffer = new MemoryStream())
                    {
                        var chunk = new byte[81920];
                        int read;
                        while ((read = await stream.ReadAsync(chunk, 0, chunk.Length, ct).ConfigureAwait(false)) > 0)
                        {
                            if (buffer.Length + read > maxBytes) return null;
                            buffer.Write(chunk, 0, read);
                        }
                        var text = Encoding.UTF8.GetString(buffer.GetBuffer(), 0, (int)buffer.Length);
                        return text.Length > 0 && text[0] == '﻿' ? text.Substring(1) : text;
                    }
                }
                catch (Exception) when (ct.IsCancellationRequested)
                {
                    throw new OperationCanceledException(ct);
                }
            }
        }

        // ": <server message>" from {"detail": "..."} or {"error": {"message": "..."}}.
        private static string DescribeServerError(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return "";
            try
            {
                var root = JObject.Parse(body);
                var message = (root["detail"] as JValue)?.Value as string
                    ?? (root["error"]?["message"] as JValue)?.Value as string;
                if (string.IsNullOrWhiteSpace(message)) return "";
                message = message.Trim();
                return ": " + (message.Length > 200 ? message.Substring(0, 200) + "…" : message);
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
}
