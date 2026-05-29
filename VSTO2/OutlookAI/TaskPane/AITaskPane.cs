using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAudio.Wave;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using OutlookAI.TaskPane.Chat;
using OutlookAI.TaskPane.Variants;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane
{
    public partial class AITaskPane : UserControl
    {
        private string _lastResult;
        private WaveInEvent _waveIn;
        private MemoryStream _pcmBuffer;
        private TextBox _activeTextBox;
        private Button _activeMicButton;
        private bool _isRecording;
        private readonly object _recordLock = new object();

        // Phase 2 (Task 30) additions
        private Outlook.Inspector _inspector;
        private LiveOutlookSurface _surface;
        private OutlookToolHost _toolHost;
        private CancellationTokenSource _activeCts;
        private Button _btnCancel;
        private Label _lblToolStrip;
        private readonly List<string> _toolStripLines = new List<string>();

        // Phase 2 (Task 34) additions
        private ConversationStore _conversationStore;
        private ChatController _chatController;

        // Phase 2 (Tasks 35-36) additions
        private VariantsController _variantsController;

        public AITaskPane()
        {
            using (TraceLog.Scope("ctor", "AITaskPane"))
            {
                InitializeComponent();
                TraceLog.Write("InitializeComponent done", "AITaskPane");
                BuildPhase2Controls();
                TraceLog.Write("BuildPhase2Controls done", "AITaskPane");
                tabControl.Selected += (s, e) =>
                    TraceLog.Write("TabControl Selected: " + (e.TabPage?.Text ?? "<null>"), "AITaskPane");
                tabControl.Deselected += (s, e) =>
                    TraceLog.Write("TabControl Deselected: " + (e.TabPage?.Text ?? "<null>"), "AITaskPane");
                this.HandleCreated += (s, e) => TraceLog.Write("HandleCreated", "AITaskPane");
                this.VisibleChanged += (s, e) => TraceLog.Write("VisibleChanged Visible=" + this.Visible, "AITaskPane");
            }
        }

        /// <summary>
        /// Bind this task pane to its owning Outlook Inspector. Called by
        /// <see cref="ThisAddIn.ShowTaskPane"/> immediately after construction.
        /// Builds the per-pane tool host so the chat service can call mailbox
        /// tools scoped to this specific compose window.
        /// </summary>
        public void Bind(Outlook.Inspector inspector)
        {
            using (TraceLog.Scope("Bind", "AITaskPane"))
            {
                _inspector = inspector;
                try
                {
                    var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                    var ids = Globals.ThisAddIn?.IdResolver;
                    var app = Globals.ThisAddIn?.Application;
                    var runner = Globals.ThisAddIn?.AdvancedSearchRunner;
                    var classifier = Globals.ThisAddIn?.FolderClassifier;
                    TraceLog.Write("Services: marshaller=" + (marshaller != null) + " ids=" + (ids != null)
                        + " app=" + (app != null) + " runner=" + (runner != null), "AITaskPane");
                    if (marshaller != null && ids != null && app != null && runner != null)
                    {
                        _surface = new LiveOutlookSurface(app, marshaller, ids,
                            composeInspector: inspector, explorer: null,
                            runner: runner, classifier: classifier);
                        _toolHost = new OutlookToolHost(_surface, Config.WriteToolsEnabled);
                        TraceLog.Write("surface + toolHost constructed", "AITaskPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("surface/toolHost error: " + ex, "AITaskPane");
                }

                // Bring up the Chat tab's WebView2 surface. Fire-and-forget; the
                // controller swaps in a fallback Label if WebView2 fails to init.
                try
                {
                    if (ChatService != null && _toolHost != null && _surface != null)
                    {
                        _conversationStore = new ConversationStore();
                        _chatController = new ChatController(
                            tabChat, ChatService, _toolHost, _surface, _conversationStore);
                        TraceLog.Write("ChatController constructed; firing InitializeAsync (fire-and-forget)", "AITaskPane");
                        var initTask = _chatController.InitializeAsync();
                        // Observe the fire-and-forget so exceptions don't get
                        // silently swallowed during this investigation.
                        initTask.ContinueWith(t =>
                        {
                            if (t.IsFaulted)
                                TraceLog.Write("ChatController.InitializeAsync FAULTED: " + t.Exception, "AITaskPane");
                            else if (t.IsCanceled)
                                TraceLog.Write("ChatController.InitializeAsync CANCELLED", "AITaskPane");
                            else
                                TraceLog.Write("ChatController.InitializeAsync completed", "AITaskPane");
                        }, TaskScheduler.Default);
                        TraceLog.Write("InitializeAsync returned (async still in flight)", "AITaskPane");
                    }
                    else
                    {
                        TraceLog.Write("ChatController NOT created (ChatService/toolHost/surface null)", "AITaskPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("ChatController construction error: " + ex, "AITaskPane");
                }

                try
                {
                    if (ChatService != null && _toolHost != null && _surface != null)
                    {
                        tabVariants.Controls.Clear();
                        _variantsController = new VariantsController(
                            tabVariants, ChatService, _toolHost, _surface,
                            insertCallback: body => InsertEmailBody(body),
                            replaceCallback: body => SetEmailBody(body));
                        TraceLog.Write("VariantsController constructed; calling BuildUi", "AITaskPane");
                        _variantsController.BuildUi();
                        TraceLog.Write("VariantsController.BuildUi returned", "AITaskPane");
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("VariantsController error: " + ex, "AITaskPane");
                }
            }
        }

        private void BuildPhase2Controls()
        {
            // Cancel button: lives next to lblStatus on the Actions tab; only
            // visible while a turn is in flight.
            _btnCancel = new Button
            {
                Text = "Cancel",
                Visible = false,
                Width = 70,
                Height = 22,
                Font = new Font("Segoe UI", 8F),
                Location = new Point(lblStatus.Location.X + 165, lblStatus.Location.Y - 3)
            };
            _btnCancel.Click += BtnCancel_Click;
            tabActions.Controls.Add(_btnCancel);

            // Tool-call strip: a single multi-line label that records each
            // tool invocation as the chat loop emits OnToolCallStart /
            // OnToolCallResult events. Sits between the status line and the
            // existing result panel.
            _lblToolStrip = new Label
            {
                Location = new Point(10, lblStatus.Location.Y + 22),
                Size = new Size(290, 60),
                AutoSize = false,
                Font = new Font("Segoe UI", 8F),
                ForeColor = Color.DarkSlateGray,
                Visible = false,
                Text = ""
            };
            tabActions.Controls.Add(_lblToolStrip);

            // Reposition panelResult down 60 px to make room for the strip.
            panelResult.Location = new Point(panelResult.Location.X, panelResult.Location.Y + 60);
        }

        private CodexChatService ChatService
            => Globals.ThisAddIn != null ? Globals.ThisAddIn.ChatService : null;

        private RealtimeVoiceService VoiceService
            => Globals.ThisAddIn != null ? Globals.ThisAddIn.VoiceService : null;

        private CodexAuthService AuthService
            => Globals.ThisAddIn != null ? Globals.ThisAddIn.AuthService : null;

        /// <summary>
        /// Call this when the task pane becomes visible for a new email.
        /// </summary>
        public void ResetForNewEmail()
        {
            txtDraftPrompt.Text = "";
            txtResult.Text = "";
            panelResult.Visible = false;
            lblStatus.Visible = false;
            _lastResult = null;
        }

        // -------------------------------------------------------------------
        // Voice capture (mic) — streams raw 16-kHz / 16-bit / mono PCM into
        // a MemoryStream, then hands that stream to RealtimeVoiceService.
        // No more on-disk WAV; no more REST POST to /v1/audio/transcriptions.
        // -------------------------------------------------------------------

        private void StartRecording(TextBox targetTextBox, Button micButton)
        {
            if (_isRecording)
            {
                StopRecordingAndTranscribe();
                return;
            }

            try
            {
                if (!RequireSignedIn("Sign in to OutlookAI before using voice input."))
                {
                    return;
                }

                _activeTextBox = targetTextBox;
                _activeMicButton = micButton;
                _pcmBuffer = new MemoryStream();

                _waveIn = new WaveInEvent
                {
                    WaveFormat = new WaveFormat(16000, 16, 1) // matches Realtime input_audio_format=pcm16
                };
                _waveIn.DataAvailable += WaveIn_DataAvailable;
                _waveIn.RecordingStopped += WaveIn_RecordingStopped;

                _isRecording = true;

                micButton.BackColor = Color.LightCoral;
                micButton.ForeColor = Color.White;
                micButton.Text = "...";
                ShowStatus("Recording... Click again to stop and transcribe.", false);

                _waveIn.StartRecording();
            }
            catch (Exception ex)
            {
                ShowStatus("Mic error: " + ex.Message, true);
                System.Diagnostics.Debug.WriteLine("Recording error: " + ex);
                CleanupRecording();
            }
        }

        private void WaveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            lock (_recordLock)
            {
                if (_pcmBuffer != null && _isRecording)
                {
                    _pcmBuffer.Write(e.Buffer, 0, e.BytesRecorded);
                }
            }
        }

        private void WaveIn_RecordingStopped(object sender, StoppedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Recording stopped event fired");
        }

        private async void StopRecordingAndTranscribe()
        {
            if (!_isRecording)
            {
                return;
            }

            _isRecording = false;
            var textBox = _activeTextBox;
            var micButton = _activeMicButton;
            byte[] pcmBytes;

            try
            {
                _waveIn?.StopRecording();
                _waveIn?.Dispose();
                _waveIn = null;

                lock (_recordLock)
                {
                    pcmBytes = _pcmBuffer != null ? _pcmBuffer.ToArray() : new byte[0];
                    _pcmBuffer?.Dispose();
                    _pcmBuffer = null;
                }

                if (micButton != null)
                {
                    micButton.BackColor = SystemColors.Control;
                    micButton.ForeColor = Color.Red;
                    micButton.Text = "\u25CF";
                }

                ShowStatus("Transcribing...", false);

                if (pcmBytes.Length < 32000) // ~1s of 16-kHz mono PCM
                {
                    ShowStatus("Recording too short. Please try again.", true);
                    return;
                }

                var voice = VoiceService;
                if (voice == null)
                {
                    ShowStatus("Voice service unavailable. Restart Outlook.", true);
                    return;
                }

                string transcription;
                using (var stream = new MemoryStream(pcmBytes, writable: false))
                {
                    transcription = await voice.TranscribeAsync(stream, CancellationToken.None);
                }

                if (!string.IsNullOrEmpty(transcription))
                {
                    InvokeOnUI(() =>
                    {
                        if (textBox != null)
                        {
                            textBox.Text = transcription;
                        }
                        ShowStatus("Transcription complete!", false);
                    });
                }
                else
                {
                    InvokeOnUI(() => ShowStatus("No speech detected. Please try again.", true));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Transcription error: " + ex);
                InvokeOnUI(() => ShowStatus("Transcription error: " + ex.Message, true));
            }
            finally
            {
                _activeTextBox = null;
                _activeMicButton = null;
            }
        }

        private void CleanupRecording()
        {
            _isRecording = false;

            lock (_recordLock)
            {
                try { _pcmBuffer?.Dispose(); } catch { /* ignore */ }
                _pcmBuffer = null;
            }

            try { _waveIn?.Dispose(); } catch { /* ignore */ }
            _waveIn = null;

            if (_activeMicButton != null)
            {
                _activeMicButton.BackColor = SystemColors.Control;
                _activeMicButton.ForeColor = Color.Red;
                _activeMicButton.Text = "\u25CF";
            }

            _activeTextBox = null;
            _activeMicButton = null;
        }

        private void btnMicDraft_Click(object sender, EventArgs e)
        {
            StartRecording(txtDraftPrompt, btnMicDraft);
        }

        // -------------------------------------------------------------------
        // Text actions — all routed through CodexChatService -> Codex backend.
        // -------------------------------------------------------------------

        private async void btnProofread_Click(object sender, EventArgs e)
        {
            await ProcessAction(CodexChatService.ActionType.Proofread);
        }

        private async void btnRevise_Click(object sender, EventArgs e)
        {
            await ProcessAction(CodexChatService.ActionType.Revise);
        }

        private async void btnShorten_Click(object sender, EventArgs e)
        {
            await ProcessAction(CodexChatService.ActionType.Shorten);
        }

        private async void btnLengthen_Click(object sender, EventArgs e)
        {
            await ProcessAction(CodexChatService.ActionType.Lengthen);
        }

        private async void btnFormal_Click(object sender, EventArgs e)
        {
            await ProcessAction(CodexChatService.ActionType.Formal);
        }

        private async void btnFriendly_Click(object sender, EventArgs e)
        {
            await ProcessAction(CodexChatService.ActionType.Friendly);
        }

        private async void btnDraft_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDraftPrompt.Text))
            {
                ShowStatus("Please enter instructions for the email you want to draft.", true);
                return;
            }
            await ProcessAction(CodexChatService.ActionType.Draft, txtDraftPrompt.Text);
        }

        private async Task ProcessAction(CodexChatService.ActionType action, string prompt = "")
        {
            TraceLog.Write(">> ProcessAction " + action, "AITaskPane");
            if (_isRecording)
            {
                CleanupRecording();
            }

            if (!RequireSignedIn("Sign in to OutlookAI before using AI actions."))
            {
                return;
            }

            var chat = ChatService;
            if (chat == null)
            {
                ShowStatus("Chat service unavailable. Restart Outlook.", true);
                return;
            }

            string emailContent = GetEmailBody();
            if (action != CodexChatService.ActionType.Draft && string.IsNullOrWhiteSpace(emailContent))
            {
                ShowStatus("No email content found. Please write something first.", true);
                return;
            }

            if (action == CodexChatService.ActionType.Draft && emailContent.Length > 4000)
            {
                emailContent = emailContent.Substring(0, 4000) + "\n[... earlier messages truncated ...]";
            }

            // Build the per-turn context. The Phase 2 system prompt is the
            // Phase 1 prompt plus a tool-awareness addendum.
            const string toolAddendum =
                "\n\nYou may call mailbox tools if you need additional context "
                + "(e.g. reading another message in the thread, searching the inbox, "
                + "or creating/categorizing follow-up drafts). Most quick edits do "
                + "not require any tools.";
            var ctx = new ConversationContext
            {
                SystemInstructions = CodexChatService.GetSystemPrompt(action) + toolAddendum,
                IncludeWriteTools = Config.WriteToolsEnabled
            };

            var userMessage = CodexChatService.BuildUserMessage(action, emailContent, prompt ?? "");
            var toolHost = _toolHost ?? new OutlookToolHost(new NullSurface(), includeWriteTools: false);
            var sink = new ActionsTabSink(this);

            _toolStripLines.Clear();
            InvokeOnUI(() =>
            {
                _lblToolStrip.Text = "";
                _lblToolStrip.Visible = false;
            });

            _activeCts = new CancellationTokenSource();
            SetUIEnabled(false);
            _btnCancel.Visible = true;
            ShowStatus("Processing...", false);

            try
            {
                var turnResult = await chat.RunTurnAsync(ctx, userMessage, toolHost, sink, _activeCts.Token);

                _lastResult = turnResult.FinalAssistantText ?? "";

                InvokeOnUI(() =>
                {
                    txtResult.Text = _lastResult;
                    panelResult.Visible = !string.IsNullOrEmpty(_lastResult);
                    string verdict;
                    switch (turnResult.StopReason)
                    {
                        case StopReason.Completed:
                            verdict = "Done! Review the result below.";
                            break;
                        case StopReason.Cancelled:
                            verdict = "Stopped. Partial result shown.";
                            break;
                        case StopReason.MaxRoundsReached:
                            verdict = "Reached max tool rounds. Partial result shown.";
                            break;
                        case StopReason.Error:
                            verdict = "Error: " + (turnResult.ErrorMessage ?? "unknown");
                            break;
                        default:
                            verdict = turnResult.StopReason.ToString();
                            break;
                    }
                    ShowStatus(verdict, turnResult.StopReason == StopReason.Error);
                    SetUIEnabled(true);
                    _btnCancel.Visible = false;
                });
            }
            catch (OperationCanceledException)
            {
                InvokeOnUI(() =>
                {
                    ShowStatus("Cancelled.", false);
                    SetUIEnabled(true);
                    _btnCancel.Visible = false;
                });
            }
            catch (Exception ex)
            {
                InvokeOnUI(() =>
                {
                    string msg = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                    ShowStatus("Error - see details.", true);
                    MessageBox.Show(msg, "OutlookAI Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    panelResult.Visible = false;
                    SetUIEnabled(true);
                    _btnCancel.Visible = false;
                });
            }
            finally
            {
                _activeCts?.Dispose();
                _activeCts = null;
                TraceLog.Write("<< ProcessAction " + action, "AITaskPane");
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            try { _activeCts?.Cancel(); } catch { /* swallowed */ }
        }

        /// <summary>
        /// Local sink that pipes tool-call events into the AITaskPane's
        /// tool-strip label. We don't stream token deltas to the strip - the
        /// final assistant text lands in the existing Result text box once
        /// the turn completes.
        /// </summary>
        private sealed class ActionsTabSink : ChatEventSink
        {
            private readonly AITaskPane _pane;
            public ActionsTabSink(AITaskPane pane) { _pane = pane; }

            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                _pane.AppendToolStrip("  ... " + name);
            }

            public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            {
                var glyph = ok ? "\u2713" : "\u26A0"; // check or warning
                _pane.AppendToolStrip("  " + glyph + " " + summary);
            }

            public override void OnError(string message)
            {
                _pane.AppendToolStrip("  ! " + (message ?? ""));
            }
        }

        private void AppendToolStrip(string line)
        {
            _toolStripLines.Add(line);
            // Keep the strip bounded so it doesn't bloat past its 60-px height.
            while (_toolStripLines.Count > 6) _toolStripLines.RemoveAt(0);
            InvokeOnUI(() =>
            {
                _lblToolStrip.Text = string.Join("\r\n", _toolStripLines);
                _lblToolStrip.Visible = _toolStripLines.Count > 0;
            });
        }

        // Fallback surface used when Bind() hasn't been called (e.g. legacy
        // task-pane creation paths that don't supply an Inspector). Returns
        // empty / null for everything so write-tools are disabled and reads
        // produce safe defaults.
        private sealed class NullSurface : IOutlookSurface
        {
            public ComposeStateResult GetCurrentComposeState(bool includeFullBody) => new ComposeStateResult();
            public IReadOnlyList<FolderResult> ListFolders() => new FolderResult[0];
            public SearchResult SearchMessages(SearchMessagesArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => new SearchResult();
            public MessageDetail ReadMessage(string messageId, bool includeFullBody) => null;
            public IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => new MessageDetail[0];
            public int CountMessages(SearchMessagesArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => 0;
            public IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => new AggregationBucket[0];
            public IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads) => new ThreadSummary[0];
            public FileSavedResult ExportExcel(ExportExcelArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => null;
            public FileSavedResult ExportPdf(ExportPdfArgs args, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken)) => null;
            public CreatedDraft CreateDraft(CreateDraftArgs args) => null;
            public CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems)
                => new CurrentSelectionResult { Folder = "", FolderId = "", Count = 0, Messages = new MessageDetail[0] };
            public void MarkAsRead(string messageId, bool read) { }
            public void FlagMessage(string messageId, string flag) { }
            public void SetCategory(string messageId, string category) { }
        }

        private void InvokeOnUI(Action action)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastResult) && InsertEmailBody(_lastResult))
            {
                panelResult.Visible = false;
                txtDraftPrompt.Text = "";
                ShowStatus("Draft inserted!", false);
            }
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastResult) && SetEmailBody(_lastResult))
            {
                panelResult.Visible = false;
                txtDraftPrompt.Text = "";
                ShowStatus("Email replaced!", false);
            }
        }

        private void btnDiscard_Click(object sender, EventArgs e)
        {
            _lastResult = null;
            panelResult.Visible = false;
            lblStatus.Visible = false;
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            using (var settingsForm = new SettingsForm())
            {
                settingsForm.ShowDialog();
            }
        }

        // -------------------------------------------------------------------
        // Outlook OOM helpers (unchanged behavior)
        // -------------------------------------------------------------------

        private string GetEmailBody()
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector != null)
                {
                    var currentItem = inspector.CurrentItem;
                    if (currentItem is Outlook.MailItem mail)
                    {
                        return mail.Body ?? "";
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetEmailBody error: " + ex.Message);
            }

            return "";
        }

        private bool InsertEmailBody(string text)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector != null)
                {
                    var currentItem = inspector.CurrentItem;
                    if (currentItem is Outlook.MailItem mail)
                    {
                        string existingBody = mail.Body ?? "";
                        mail.Body = text + "\n\n" + existingBody;
                        return true;
                    }
                }
                ShowStatus("Could not find active email window.", true);
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("InsertEmailBody error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
        }

        private bool SetEmailBody(string text)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector != null)
                {
                    var currentItem = inspector.CurrentItem;
                    if (currentItem is Outlook.MailItem mail)
                    {
                        mail.Body = text;
                        return true;
                    }
                }
                ShowStatus("Could not find active email window.", true);
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SetEmailBody error: " + ex.Message);
                ShowStatus("Could not update email: " + ex.Message, true);
                return false;
            }
        }

        private void ShowStatus(string message, bool isError)
        {
            lblStatus.Text = message;
            lblStatus.ForeColor = isError ? Color.DarkRed : Color.DarkGreen;
            lblStatus.Visible = true;
        }

        private void SetUIEnabled(bool enabled)
        {
            btnProofread.Enabled = enabled;
            btnRevise.Enabled = enabled;
            btnShorten.Enabled = enabled;
            btnLengthen.Enabled = enabled;
            btnFormal.Enabled = enabled;
            btnFriendly.Enabled = enabled;
            btnDraft.Enabled = enabled;
            btnMicDraft.Enabled = enabled;
            txtDraftPrompt.Enabled = enabled;
        }

        private bool RequireSignedIn(string promptMessage)
        {
            var auth = AuthService;
            if (auth == null)
            {
                ShowStatus("Auth service unavailable. Restart Outlook.", true);
                return false;
            }
            var status = auth.GetStatus();
            if (status.State == AuthState.Authenticated)
            {
                return true;
            }
            ShowStatus(promptMessage, true);
            using (var settingsForm = new SettingsForm(auth))
            {
                settingsForm.ShowDialog();
            }
            return auth.GetStatus().State == AuthState.Authenticated;
        }

        partial void DisposeCustomResources()
        {
            CleanupRecording();
            try { _chatController?.Dispose(); } catch { }
            try { _variantsController?.Dispose(); } catch { }
        }
    }
}
