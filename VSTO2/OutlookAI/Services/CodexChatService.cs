using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    /// <summary>
    /// ChatGPT consumer-subscription text generation via the Codex Responses
    /// backend. Replaces the legacy Anthropic <c>ClaudeService</c>.
    ///
    /// Endpoint: <c>https://chatgpt.com/backend-api/codex/responses</c>.
    /// Auth: <c>Authorization: Bearer &lt;OAuth access_token&gt;</c> from
    /// <see cref="CodexAuthService.GetAccessTokenAsync"/>.
    /// Wire format: Codex Responses API request shape; SSE response stream.
    /// </summary>
    public sealed class CodexChatService : IDisposable
    {
        public const string ResponsesEndpoint = "https://chatgpt.com/backend-api/codex/responses";

        public enum ActionType
        {
            Proofread,
            Revise,
            Draft,
            Shorten,
            Lengthen,
            Formal,
            Friendly,
            Custom
        }

        private readonly CodexAuthService _auth;
        private readonly HttpClient _http;
        private readonly bool _ownsHttp;
        private bool _disposed;

        public CodexChatService(CodexAuthService auth)
            : this(auth, BuildDefaultHttpClient(), ownsHttp: true)
        {
        }

        // Test seam.
        public CodexChatService(CodexAuthService auth, HttpClient httpClient, bool ownsHttp = false)
        {
            _auth = auth ?? throw new ArgumentNullException(nameof(auth));
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _ownsHttp = ownsHttp;
        }

        private static HttpClient BuildDefaultHttpClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            return new HttpClient
            {
                // Long timeout because the SSE stream is held open until the
                // model finishes generating.
                Timeout = TimeSpan.FromMinutes(2)
            };
        }

        public async Task<string> ProcessEmailAsync(
            ActionType action,
            string emailContent,
            string customPrompt = "",
            CancellationToken cancellationToken = default(CancellationToken))
        {
            var instructions = GetSystemPrompt(action);
            var userMessage = BuildUserMessage(action, emailContent, customPrompt ?? "");
            var body = BuildResponsesRequest(instructions, userMessage);

            using (var request = new HttpRequestMessage(HttpMethod.Post, ResponsesEndpoint))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue(
                    "Bearer",
                    await _auth.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false));
                request.Headers.Accept.ParseAdd("text/event-stream");

                var status = _auth.GetStatus();
                if (status.State == AuthState.Authenticated && !string.IsNullOrEmpty(status.AccountId))
                {
                    request.Headers.TryAddWithoutValidation("ChatGPT-Account-ID", status.AccountId);
                }

                request.Content = new StringContent(body.ToString(Newtonsoft.Json.Formatting.None), Encoding.UTF8, "application/json");

                using (var response = await _http.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        var errorBody = await SafeReadAsStringAsync(response).ConfigureAwait(false);
                        throw new InvalidOperationException(
                            "ChatGPT Codex backend error: " + (int)response.StatusCode + " " + errorBody);
                    }

                    using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    {
                        return await ReadResponsesSseAsync(stream, cancellationToken).ConfigureAwait(false);
                    }
                }
            }
        }

        private static JObject BuildResponsesRequest(string instructions, string userMessage)
        {
            // Mirrors codex-rs/codex-api/src/common.rs::ResponsesApiRequest.
            return new JObject(
                new JProperty("model", Config.Model),
                new JProperty("instructions", instructions ?? ""),
                new JProperty("input", new JArray(
                    new JObject(
                        new JProperty("type", "message"),
                        new JProperty("role", "user"),
                        new JProperty("content", new JArray(
                            new JObject(
                                new JProperty("type", "input_text"),
                                new JProperty("text", userMessage ?? ""))))))),
                new JProperty("tools", new JArray()),
                new JProperty("tool_choice", "auto"),
                new JProperty("parallel_tool_calls", false),
                new JProperty("reasoning", JValue.CreateNull()),
                new JProperty("store", false),
                new JProperty("stream", true),
                new JProperty("include", new JArray()));
        }

        private static async Task<string> SafeReadAsStringAsync(HttpResponseMessage response)
        {
            try
            {
                return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch
            {
                return "";
            }
        }

        // Reads the Responses SSE stream and accumulates output text deltas.
        // The final assistant text arrives either as a sequence of
        // `response.output_text.delta` events ending with `response.completed`,
        // or as a single `response.output_text.done` event with `text` set.
        private static async Task<string> ReadResponsesSseAsync(Stream stream, CancellationToken cancellationToken)
        {
            var output = new StringBuilder();
            using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, bufferSize: 8 * 1024, leaveOpen: false))
            {
                while (!reader.EndOfStream)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var line = await reader.ReadLineAsync().ConfigureAwait(false);
                    if (line == null)
                    {
                        break;
                    }
                    if (line.Length == 0 || !line.StartsWith("data:", StringComparison.Ordinal))
                    {
                        continue;
                    }
                    var payload = line.Substring(5).TrimStart();
                    if (payload == "[DONE]")
                    {
                        break;
                    }

                    JObject evt;
                    try
                    {
                        evt = JObject.Parse(payload);
                    }
                    catch
                    {
                        continue;
                    }

                    var type = (string)evt["type"];
                    if (string.IsNullOrEmpty(type))
                    {
                        continue;
                    }

                    if (type == "response.output_text.delta")
                    {
                        var delta = (string)evt["delta"];
                        if (!string.IsNullOrEmpty(delta))
                        {
                            output.Append(delta);
                        }
                    }
                    else if (type == "response.output_text.done")
                    {
                        var text = (string)evt["text"];
                        if (!string.IsNullOrEmpty(text))
                        {
                            // Replace accumulated deltas with the final, canonical text.
                            output.Length = 0;
                            output.Append(text);
                        }
                    }
                    else if (type == "response.completed")
                    {
                        break;
                    }
                    else if (type == "error")
                    {
                        var err = evt["error"] as JObject;
                        var message = err != null ? (string)err["message"] : null;
                        throw new InvalidOperationException("ChatGPT Codex backend error: " + (message ?? payload));
                    }
                }
            }

            return output.ToString().Trim();
        }

        // -------------------------------------------------------------------
        // Prompt + message construction (preserves v1 ClaudeService behavior)
        // -------------------------------------------------------------------

        public static string GetSystemPrompt(ActionType action)
        {
            switch (action)
            {
                case ActionType.Proofread:
                    return "You are a professional editor. Review the email for grammar, spelling, punctuation, and clarity issues. Return the corrected email text only. Do not add any explanations.";
                case ActionType.Revise:
                    return "You are a professional writing assistant. Improve the email clarity, flow, and impact. Return only the revised email text without any explanations.";
                case ActionType.Draft:
                    return "You are a professional email writer. Write a clear, professional email based on the instructions. If replying to an email thread, write only your reply - do not include the previous messages. Return only the email text you are composing.";
                case ActionType.Shorten:
                    return "You are a professional editor. Condense this email to be more concise while keeping essential information. Return only the shortened email text.";
                case ActionType.Lengthen:
                    return "You are a professional writer. Expand this email with more detail while maintaining professionalism. Return only the expanded email text.";
                case ActionType.Formal:
                    return "You are a professional editor. Rewrite this email in a more formal tone suitable for business. Return only the rewritten email text.";
                case ActionType.Friendly:
                    return "You are a professional editor. Rewrite this email in a warmer, friendlier tone while remaining professional. Return only the rewritten email text.";
                case ActionType.Custom:
                default:
                    return "You are a professional email writing assistant. Help the user with their email based on their instructions. Return only the result.";
            }
        }

        public static string BuildUserMessage(ActionType action, string emailContent, string customPrompt)
        {
            if (action == ActionType.Draft)
            {
                if (!string.IsNullOrWhiteSpace(emailContent))
                {
                    return "Write a reply email based on these instructions:\n\n" + customPrompt +
                           "\n\n--- Email thread for context (do NOT include this in your response, just use it for context) ---\n\n" + emailContent;
                }
                return "Write an email based on these instructions:\n\n" + customPrompt;
            }
            if (action == ActionType.Custom)
            {
                return "Email content:\n\n" + emailContent + "\n\nInstructions: " + customPrompt;
            }
            return "Email to " + action.ToString().ToLowerInvariant() + ":\n\n" + emailContent;
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
        }
    }
}
