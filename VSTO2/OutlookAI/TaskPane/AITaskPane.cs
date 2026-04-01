using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Text;
using NAudio.Wave;
using OutlookAI.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.TaskPane
{
    public partial class AITaskPane : UserControl
    {
        private readonly ClaudeService _claudeService;
        private string _lastResult;
        private WaveInEvent _waveIn;
        private WaveFileWriter _waveWriter;
        private string _tempAudioFile;
        private TextBox _activeTextBox;
        private Button _activeMicButton;
        private bool _isRecording = false;
        private readonly object _recordLock = new object();

        public AITaskPane()
        {
            InitializeComponent();
            _claudeService = new ClaudeService();
        }

        /// <summary>
        /// Call this when the task pane becomes visible for a new email
        /// </summary>
        public void ResetForNewEmail()
        {
            txtDraftPrompt.Text = "";
            txtResult.Text = "";
            panelResult.Visible = false;
            lblStatus.Visible = false;
            _lastResult = null;
        }

        private void StartRecording(TextBox targetTextBox, Button micButton)
        {
            if (_isRecording)
            {
                StopRecordingAndTranscribe();
                return;
            }

            try
            {
                _activeTextBox = targetTextBox;
                _activeMicButton = micButton;

                // Create temp file for audio
                _tempAudioFile = Path.Combine(Path.GetTempPath(), $"outlook_ai_{Guid.NewGuid()}.wav");

                // Setup audio recording
                _waveIn = new WaveInEvent();
                _waveIn.WaveFormat = new WaveFormat(16000, 16, 1); // 16kHz, 16-bit, mono (optimal for Whisper)
                _waveIn.DataAvailable += WaveIn_DataAvailable;
                _waveIn.RecordingStopped += WaveIn_RecordingStopped;

                _waveWriter = new WaveFileWriter(_tempAudioFile, _waveIn.WaveFormat);

                _isRecording = true;

                // Visual feedback
                micButton.BackColor = Color.LightCoral;
                micButton.ForeColor = Color.White;
                micButton.Text = "...";
                ShowStatus("Recording... Click again to stop and transcribe.", false);

                _waveIn.StartRecording();
                System.Diagnostics.Debug.WriteLine("Recording started: " + _tempAudioFile);
            }
            catch (Exception ex)
            {
                ShowStatus("Mic error: " + ex.Message, true);
                System.Diagnostics.Debug.WriteLine("Recording error: " + ex.Message);
                CleanupRecording();
            }
        }

        private void WaveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            lock (_recordLock)
            {
                if (_waveWriter != null && _isRecording)
                {
                    _waveWriter.Write(e.Buffer, 0, e.BytesRecorded);
                }
            }
        }

        private void WaveIn_RecordingStopped(object sender, StoppedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Recording stopped event fired");
        }

        private async void StopRecordingAndTranscribe()
        {
            if (!_isRecording) return;

            _isRecording = false;
            var textBox = _activeTextBox;
            var micButton = _activeMicButton;
            var audioFile = _tempAudioFile;

            try
            {
                // Stop recording
                _waveIn?.StopRecording();

                lock (_recordLock)
                {
                    _waveWriter?.Dispose();
                    _waveWriter = null;
                }

                _waveIn?.Dispose();
                _waveIn = null;

                // Update UI
                if (micButton != null)
                {
                    micButton.BackColor = SystemColors.Control;
                    micButton.ForeColor = Color.Red;
                    micButton.Text = "\u25CF";
                }

                ShowStatus("Transcribing...", false);

                // Check file size
                var fileInfo = new FileInfo(audioFile);
                System.Diagnostics.Debug.WriteLine($"Audio file size: {fileInfo.Length} bytes");

                if (fileInfo.Length < 1000)
                {
                    ShowStatus("Recording too short. Please try again.", true);
                    return;
                }

                // Send to Whisper API
                string transcription = await Task.Run(() => TranscribeWithWhisper(audioFile));

                if (!string.IsNullOrEmpty(transcription))
                {
                    InvokeOnUI(() =>
                    {
                        if (textBox != null)
                        {
                            // Replace text instead of appending
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
                System.Diagnostics.Debug.WriteLine("Transcription error: " + ex.Message);
                InvokeOnUI(() => ShowStatus("Transcription error: " + ex.Message, true));
            }
            finally
            {
                // Cleanup temp file
                try
                {
                    if (File.Exists(audioFile))
                        File.Delete(audioFile);
                }
                catch { }

                _activeTextBox = null;
                _activeMicButton = null;
            }
        }

        private string TranscribeWithWhisper(string audioFilePath)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string boundary = "----WebKitFormBoundary" + DateTime.Now.Ticks.ToString("x");
                byte[] fileBytes = File.ReadAllBytes(audioFilePath);

                // Build multipart body in memory
                using (var bodyStream = new MemoryStream())
                {
                    var encoding = new UTF8Encoding(false); // No BOM

                    // File field
                    byte[] fileHeader = encoding.GetBytes(
                        $"--{boundary}\r\n" +
                        $"Content-Disposition: form-data; name=\"file\"; filename=\"audio.wav\"\r\n" +
                        $"Content-Type: audio/wav\r\n\r\n");
                    bodyStream.Write(fileHeader, 0, fileHeader.Length);
                    bodyStream.Write(fileBytes, 0, fileBytes.Length);

                    // Model field
                    byte[] modelField = encoding.GetBytes(
                        $"\r\n--{boundary}\r\n" +
                        $"Content-Disposition: form-data; name=\"model\"\r\n\r\n" +
                        $"{Config.WhisperModel}");
                    bodyStream.Write(modelField, 0, modelField.Length);

                    // Language field
                    byte[] langField = encoding.GetBytes(
                        $"\r\n--{boundary}\r\n" +
                        $"Content-Disposition: form-data; name=\"language\"\r\n\r\n" +
                        $"en");
                    bodyStream.Write(langField, 0, langField.Length);

                    // Closing boundary
                    byte[] closing = encoding.GetBytes($"\r\n--{boundary}--\r\n");
                    bodyStream.Write(closing, 0, closing.Length);

                    byte[] bodyBytes = bodyStream.ToArray();

                    var request = (HttpWebRequest)WebRequest.Create("https://api.openai.com/v1/audio/transcriptions");
                    request.Method = "POST";
                    request.ContentType = "multipart/form-data; boundary=" + boundary;
                    request.ContentLength = bodyBytes.Length;
                    request.Headers.Add("Authorization", "Bearer " + Config.OpenAIApiKey);

                    using (var requestStream = request.GetRequestStream())
                    {
                        requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                    }

                    using (var response = (HttpWebResponse)request.GetResponse())
                    using (var reader = new StreamReader(response.GetResponseStream()))
                    {
                        string json = reader.ReadToEnd();
                        System.Diagnostics.Debug.WriteLine("Whisper response: " + json);

                        // Parse "text" field from JSON
                        int textStart = json.IndexOf("\"text\":");
                        if (textStart >= 0)
                        {
                            textStart = json.IndexOf("\"", textStart + 7) + 1;
                            int textEnd = json.IndexOf("\"", textStart);
                            if (textEnd > textStart)
                            {
                                string text = json.Substring(textStart, textEnd - textStart);
                                return text.Replace("\\n", "\n").Replace("\\\"", "\"");
                            }
                        }
                    }
                }
            }
            catch (WebException ex)
            {
                if (ex.Response != null)
                {
                    using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        string error = reader.ReadToEnd();
                        System.Diagnostics.Debug.WriteLine("Whisper API error: " + error);
                        throw new Exception("Whisper API: " + error);
                    }
                }
                throw;
            }

            return null;
        }

        private void CleanupRecording()
        {
            _isRecording = false;

            lock (_recordLock)
            {
                try { _waveWriter?.Dispose(); } catch { }
                _waveWriter = null;
            }

            try { _waveIn?.Dispose(); } catch { }
            _waveIn = null;

            if (_activeMicButton != null)
            {
                _activeMicButton.BackColor = SystemColors.Control;
                _activeMicButton.ForeColor = Color.Red;
                _activeMicButton.Text = "\u25CF";
            }

            try
            {
                if (!string.IsNullOrEmpty(_tempAudioFile) && File.Exists(_tempAudioFile))
                    File.Delete(_tempAudioFile);
            }
            catch { }

            _activeTextBox = null;
            _activeMicButton = null;
        }

        private void btnMicDraft_Click(object sender, EventArgs e)
        {
            StartRecording(txtDraftPrompt, btnMicDraft);
        }

        // SetMailItem removed - we always use ActiveInspector now

        private async void btnProofread_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Proofread);
        }

        private async void btnRevise_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Revise);
        }

        private async void btnShorten_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Shorten);
        }

        private async void btnLengthen_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Lengthen);
        }

        private async void btnFormal_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Formal);
        }

        private async void btnFriendly_Click(object sender, EventArgs e)
        {
            await ProcessAction(ClaudeService.ActionType.Friendly);
        }

        private async void btnDraft_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDraftPrompt.Text))
            {
                ShowStatus("Please enter instructions for the email you want to draft.", true);
                return;
            }
            await ProcessAction(ClaudeService.ActionType.Draft, txtDraftPrompt.Text);
        }

        private async Task ProcessAction(ClaudeService.ActionType action, string prompt = "")
        {
            // Stop any active recording
            if (_isRecording) CleanupRecording();

            string emailContent = GetEmailBody();

            // For non-Draft actions, we need existing content to work with
            if (action != ClaudeService.ActionType.Draft && string.IsNullOrWhiteSpace(emailContent))
            {
                ShowStatus("No email content found. Please write something first.", true);
                return;
            }

            // For Draft, truncate chain to avoid token limits (keep ~4000 chars)
            if (action == ClaudeService.ActionType.Draft && emailContent.Length > 4000)
            {
                emailContent = emailContent.Substring(0, 4000) + "\n[... earlier messages truncated ...]";
            }

            SetUIEnabled(false);
            ShowStatus("Processing...", false);

            try
            {
                string result = await Task.Run(() =>
                    _claudeService.ProcessEmailAsync(action, emailContent, prompt).Result);

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
            if (!string.IsNullOrEmpty(_lastResult))
            {
                if (InsertEmailBody(_lastResult))
                {
                    panelResult.Visible = false;
                    txtDraftPrompt.Text = "";
                    ShowStatus("Draft inserted!", false);
                }
            }
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastResult))
            {
                if (SetEmailBody(_lastResult))
                {
                    panelResult.Visible = false;
                    txtDraftPrompt.Text = "";
                    ShowStatus("Email replaced!", false);
                }
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
                        // Insert at top with separator
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

        partial void DisposeCustomResources()
        {
            CleanupRecording();
        }
    }

    public class SettingsForm : Form
    {
        private TextBox txtPassword;
        private TextBox txtApiKey;
        private ComboBox cboModel;
        private NumericUpDown numMaxTokens;
        private TextBox txtNewPassword;
        private Button btnSave;
        private Panel panelSettings;
        private Label lblError;
        private bool _authenticated = false;

        public SettingsForm()
        {
            this.Text = "AI Assistant Settings";
            this.Size = new Size(400, 350);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lblPassword = new Label { Text = "Admin Password:", Location = new Point(20, 20), AutoSize = true };
            txtPassword = new TextBox { Location = new Point(20, 45), Width = 340, PasswordChar = '*' };
            var btnLogin = new Button { Text = "Login", Location = new Point(280, 75), Width = 80 };
            btnLogin.Click += BtnLogin_Click;

            lblError = new Label { Location = new Point(20, 80), AutoSize = true, ForeColor = Color.DarkRed, Visible = false };

            panelSettings = new Panel { Location = new Point(0, 110), Size = new Size(400, 200), Visible = false };

            var lblApiKey = new Label { Text = "API Key (leave blank to keep current):", Location = new Point(20, 10), AutoSize = true };
            txtApiKey = new TextBox { Location = new Point(20, 30), Width = 340, PasswordChar = '*' };

            var lblModel = new Label { Text = "Model:", Location = new Point(20, 60), AutoSize = true };
            cboModel = new ComboBox { Location = new Point(20, 80), Width = 340, DropDownStyle = ComboBoxStyle.DropDownList };
            cboModel.Items.AddRange(Config.AvailableModels);
            cboModel.SelectedItem = Config.Model;

            var lblMaxTokens = new Label { Text = "Max Tokens:", Location = new Point(20, 110), AutoSize = true };
            numMaxTokens = new NumericUpDown { Location = new Point(20, 130), Width = 100, Minimum = 256, Maximum = 4096, Value = Config.MaxTokens };

            var lblNewPassword = new Label { Text = "New Password (leave blank to keep):", Location = new Point(20, 160), AutoSize = true };
            txtNewPassword = new TextBox { Location = new Point(20, 180), Width = 200, PasswordChar = '*' };

            btnSave = new Button { Text = "Save", Location = new Point(200, 210), Width = 80 };
            btnSave.Click += BtnSave_Click;

            var btnCancel = new Button { Text = "Cancel", Location = new Point(290, 210), Width = 80 };
            btnCancel.Click += (s, e) => this.Close();

            panelSettings.Controls.AddRange(new Control[] { lblApiKey, txtApiKey, lblModel, cboModel, lblMaxTokens, numMaxTokens, lblNewPassword, txtNewPassword, btnSave, btnCancel });
            this.Controls.AddRange(new Control[] { lblPassword, txtPassword, btnLogin, lblError, panelSettings });
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            if (txtPassword.Text == Config.AdminPassword)
            {
                _authenticated = true;
                panelSettings.Visible = true;
                lblError.Visible = false;
                txtPassword.Enabled = false;
            }
            else
            {
                lblError.Text = "Invalid password";
                lblError.Visible = true;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!_authenticated) return;
            if (!string.IsNullOrWhiteSpace(txtApiKey.Text)) Config.ApiKey = txtApiKey.Text;
            if (cboModel.SelectedItem != null) Config.Model = cboModel.SelectedItem.ToString();
            Config.MaxTokens = (int)numMaxTokens.Value;
            if (!string.IsNullOrWhiteSpace(txtNewPassword.Text)) Config.AdminPassword = txtNewPassword.Text;
            Config.SaveConfig();
            MessageBox.Show("Settings saved!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}
