using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;

namespace OutlookAI.TaskPane.Chat
{
    /// <summary>
    /// Owns the Chat tab's WebView2 lifecycle and the JS&#x2194;C# bridge that
    /// wires user input into <see cref="CodexChatService.RunTurnAsync"/>.
    /// Per-Inspector instance, constructed by <see cref="AITaskPane"/> after
    /// <see cref="AITaskPane.Bind"/> hands it the tool host + surface.
    /// </summary>
    public sealed class ChatController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly CodexChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly Func<string> _composerSystemPrompt;

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public ChatController(
            Control hostContainer,
            CodexChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            _composerSystemPrompt = () =>
                "You are an AI assistant embedded in Microsoft Outlook's compose window. "
                + "Help the user understand, draft, and revise the email in front of them. "
                + "You have mailbox tools available for context (read other messages, search, "
                + "list folders). Prefer one focused tool call over many. Reply concisely; "
                + "the user is busy.";
        }

        /// <summary>
        /// Initialize the WebView2, extract WebUI resources, wire bridge.
        /// Awaited by <see cref="AITaskPane.Bind"/>. On failure, leaves a
        /// friendly fallback label on the tab.
        /// </summary>
        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "ChatController");
            if (!WebView2Bootstrap.IsRuntimeInstalled())
            {
                TraceLog.Write("WebView2 runtime NOT installed; showing fallback", "ChatController");
                ShowFallback("WebView2 runtime not installed.\r\nRun the installer or download:\r\n" +
                             "https://developer.microsoft.com/microsoft-edge/webview2/");
                return;
            }
            TraceLog.Write("WebView2 runtime detected; constructing WebView2 control", "ChatController");

            _webView = new WebView2 { Dock = DockStyle.Fill };
            TraceLog.Write("WebView2 control constructed", "ChatController");
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_webView);
            TraceLog.Write("WebView2 control added to host container; about to await Bootstrap.InitializeAsync", "ChatController");

            try
            {
                await WebView2Bootstrap.InitializeAsync(_webView);
                TraceLog.Write("Bootstrap.InitializeAsync awaited OK", "ChatController");
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                TraceLog.Write("WebMessageReceived subscribed; navigating", "ChatController");
                _webView.CoreWebView2.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
                TraceLog.Write("Navigate called", "ChatController");
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("ChatController.InitializeAsync: " + ex);
                ShowFallback("WebView2 failed to initialize: " + ex.Message);
            }
        }

        private void ShowFallback(string message)
        {
            if (_fallbackLabel != null)
            {
                _fallbackLabel.Text = message;
                return;
            }
            _fallbackLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkSlateGray,
                Text = message
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_fallbackLabel);
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                TraceLog.Write("WebMessageReceived: " + (json?.Length > 80 ? json.Substring(0, 80) + "..." : json), "ChatController");
                if (string.IsNullOrEmpty(json)) return;
                var obj = JObject.Parse(json);
                var type = (string)obj["type"] ?? "";
                var payload = obj["payload"] as JObject;
                HandleHostMessage(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("ChatController bridge parse error: " + ex);
            }
        }

        private void HandleHostMessage(string type, JObject payload)
        {
            switch (type)
            {
                case "ready":
                    OnWebViewReady();
                    break;
                case "send":
                    _ = StartTurnAsync(
                        (string)payload?["text"] ?? "",
                        (string)payload?["reasoning"]);
                    break;
                case "stop":
                    try { _activeCts?.Cancel(); } catch { }
                    break;
                case "clear":
                    _store.Clear();
                    _ = RunScript("outlookai.clear();");
                    break;
                case "copy":
                    var clip = _store.ExportForClipboard();
                    try { Clipboard.SetText(clip ?? ""); } catch { /* clipboard occasionally throws on Outlook */ }
                    break;
            }
        }

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "ChatController");
            _isReady = true;
            _ = RunScript("outlookai.applyTheme('light');");
            PushContextStripFromSurface();
            TraceLog.Write("OnWebViewReady completed", "ChatController");
        }

        private void PushContextStripFromSurface()
        {
            if (_surface == null) { TraceLog.Write("PushContextStrip: _surface is null", "ChatController"); return; }
            try
            {
                TraceLog.Write(">> PushContextStrip calling GetCurrentComposeState", "ChatController");
                var state = _surface.GetCurrentComposeState(includeFullBody: false);
                TraceLog.Write("<< PushContextStrip GetCurrentComposeState returned", "ChatController");
                var ctx = new JObject(
                    new JProperty("subject", state?.Subject ?? ""),
                    new JProperty("recipients", new JArray((state?.ToRecipients ?? new string[0]).Take(3))),
                    new JProperty("thread", state?.InReplyTo?.ThreadTopic ?? ""));
                _ = RunScript("outlookai.setContextStrip(" + ctx.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushContextStrip EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("PushContextStrip error: " + ex);
            }
        }

        private async Task StartTurnAsync(string userText, string reasoningOverride)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "ChatController");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "ChatController");
                return;
            }
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();

            await RunScript("outlookai.appendUserMessage(" + JsString(userText) + ");");
            await RunScript("outlookai.setComposerEnabled(false, true);");
            var assistantId = "asst_" + (++_nextMessageId);
            await RunScript("outlookai.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                // RunTurnAsync mutates context.History in place (adds the user
                // message, then assistant/function-call/output items). We start
                // with a snapshot of the existing store, let the turn evolve it,
                // and then sync the diff back into the store at the end.
                var initialSnapshot = _store.Snapshot();
                var ctx = new ConversationContext
                {
                    SystemInstructions = BuildSystemInstructionsWithComposeContext(),
                    History = new System.Collections.Generic.List<JObject>(initialSnapshot),
                    IncludeWriteTools = Config.WriteToolsEnabled,
                    ReasoningEffortOverride = string.IsNullOrEmpty(reasoningOverride) ? null : reasoningOverride
                };

                var sink = new WebViewSink(this, assistantId);
                var result = await _chat.RunTurnAsync(ctx, userText, _toolHost, sink, _activeCts.Token);

                // Sync newly-appended items (user msg + assistant + tool round-
                // trips) back into the store so the next turn starts from the
                // updated history.
                for (int i = initialSnapshot.Count; i < ctx.History.Count; i++)
                {
                    _store.Append(ctx.History[i]);
                }

                var opts = new JObject(
                    new JProperty("stopped", result.StopReason == StopReason.Cancelled),
                    new JProperty("error", result.StopReason == StopReason.Error));
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", " +
                                opts.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (OperationCanceledException)
            {
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) +
                                ", {stopped:true});");
            }
            catch (Exception ex)
            {
                await RunScript("outlookai.showError(" + JsString(ex.Message ?? "") + ");");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("outlookai.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartTurnAsync", "ChatController");
            }
        }

        private string BuildSystemInstructionsWithComposeContext()
        {
            var prompt = _composerSystemPrompt();
            try
            {
                var state = _surface?.GetCurrentComposeState(includeFullBody: false);
                if (state != null)
                {
                    var sb = new System.Text.StringBuilder(prompt);
                    sb.AppendLine();
                    sb.AppendLine();
                    sb.AppendLine("---");
                    sb.AppendLine("Current compose state (read-only context):");
                    if (!string.IsNullOrEmpty(state.Subject))
                    {
                        sb.AppendLine("Subject: " + state.Subject);
                    }
                    if (state.ToRecipients != null && state.ToRecipients.Any())
                    {
                        sb.AppendLine("To: " + string.Join(", ", state.ToRecipients));
                    }
                    if (state.CcRecipients != null && state.CcRecipients.Any())
                    {
                        sb.AppendLine("Cc: " + string.Join(", ", state.CcRecipients));
                    }
                    if (state.InReplyTo != null && !string.IsNullOrEmpty(state.InReplyTo.ThreadTopic))
                    {
                        sb.AppendLine("Reply-to thread: " + state.InReplyTo.ThreadTopic);
                    }
                    if (!string.IsNullOrEmpty(state.BodyPlaintext))
                    {
                        sb.AppendLine("Body (current draft, may be empty):");
                        sb.AppendLine(state.BodyPlaintext);
                    }
                    sb.AppendLine("---");
                    return sb.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("BuildSystemInstructions compose-state error: " + ex);
            }
            return prompt;
        }

        private async Task RunScript(string script)
        {
            if (_isDisposed) return;
            try
            {
                // EVERY WebView2 property/method access (including just
                // reading _webView.CoreWebView2) MUST happen on the UI thread.
                // Marshal first, THEN touch the control.
                var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                if (marshaller != null
                    && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
                {
                    await marshaller.RunAsync(() =>
                    {
                        if (_isDisposed) return;
                        var core = _webView?.CoreWebView2;
                        if (core == null) return;
                        _ = core.ExecuteScriptAsync(script);
                    }, CancellationToken.None).ConfigureAwait(false);
                    return;
                }
                // On UI thread: safe to touch the control directly.
                var coreOnUi = _webView?.CoreWebView2;
                if (coreOnUi == null) return;
                await coreOnUi.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "ChatController");
                System.Diagnostics.Debug.WriteLine("ExecuteScriptAsync failed: " + ex);
            }
        }

        // JSON-encode a string for safe inlining inside an ExecuteScriptAsync
        // call. Newtonsoft handles quote / backslash / control characters.
        private static string JsString(string s)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
            try { _webView?.Dispose(); } catch { }
        }

        /// <summary>
        /// Streams ChatEventSink callbacks back into the WebView2 surface via
        /// ExecuteScriptAsync. Fire-and-forget; failures are swallowed and
        /// logged because the chat loop should not stall on UI hiccups.
        /// </summary>
        private sealed class WebViewSink : ChatEventSink
        {
            private readonly ChatController _owner;
            private readonly string _assistantId;
            public WebViewSink(ChatController owner, string assistantId)
            {
                _owner = owner;
                _assistantId = assistantId;
            }
            public override void OnTokenDelta(string delta)
            {
                TraceLog.Write("Sink.OnTokenDelta len=" + (delta?.Length ?? 0), "WebViewSink");
                _ = _owner.RunScript(
                    "outlookai.appendTextDelta(" +
                    JsString(_assistantId) + ", " + JsString(delta) + ");");
            }
            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                TraceLog.Write("Sink.OnToolCallStart " + name + " id=" + callId, "WebViewSink");
                _ = _owner.RunScript(
                    "outlookai.appendToolCallCard(" +
                    JsString(callId) + ", " + JsString(name) + ", " + JsString(argsJson) + ");");
            }
            public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            {
                TraceLog.Write("Sink.OnToolCallResult ok=" + ok + " id=" + callId, "WebViewSink");
                _ = _owner.RunScript(
                    "outlookai.updateToolCallCard(" +
                    JsString(callId) + ", " + (ok ? "true" : "false") + ", " +
                    JsString(summary) + ", " + JsString(resultJson) + ");");
                if (IsWriteTool(callId) && ok)
                {
                    _ = _owner.RunScript(
                        "outlookai.appendAuditRow(" + JsString("Wrote: " + summary) + ");");
                }
            }
            public override void OnError(string message)
            {
                TraceLog.Write("Sink.OnError: " + message, "WebViewSink");
                _ = _owner.RunScript("outlookai.showError(" + JsString(message ?? "") + ");");
            }
            public override void OnAssistantMessageComplete(string text)
            {
                TraceLog.Write("Sink.OnAssistantMessageComplete len=" + (text?.Length ?? 0), "WebViewSink");
            }
            public override void OnRoundBoundary()
            {
                TraceLog.Write("Sink.OnRoundBoundary", "WebViewSink");
            }

            // The C# side doesn't currently track which callIds correspond
            // to write tools - the model decides at runtime. For now we
            // detect write-tool intent by name prefix as a best effort.
            private static bool IsWriteTool(string callId)
            {
                // CallId doesn't carry the tool name; this method is a stub
                // until we plumb the name through. Returning false keeps the
                // audit row off until the explicit plumbing is in place
                // (avoids false-positive audit rows for read tools).
                return false;
            }
        }
    }
}
