using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAudio.Wave;
using OutlookAI.Services;
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

        public AITaskPane()
        {
            InitializeComponent();
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

            // For Draft, truncate the chain so the prompt doesn't blow past the model context.
            if (action == CodexChatService.ActionType.Draft && emailContent.Length > 4000)
            {
                emailContent = emailContent.Substring(0, 4000) + "\n[... earlier messages truncated ...]";
            }

            SetUIEnabled(false);
            ShowStatus("Processing...", false);

            try
            {
                string result = await chat.ProcessEmailAsync(action, emailContent, prompt, CancellationToken.None);
                _lastResult = result;

                InvokeOnUI(() =>
                {
                    txtResult.Text = _lastResult;
                    panelResult.Visible = true;
                    ShowStatus("Done! Review the result below.", false);
                    SetUIEnabled(true);
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
                });
            }
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
        }
    }
}
