using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace OutlookAI.QuickTools
{
    internal sealed class QuickToolsProgressForm : Form
    {
        private readonly Label _status;
        private readonly Button _cancel;
        private readonly CancellationTokenSource _cts;
        private bool _finished;

        public QuickToolsProgressForm(CancellationTokenSource cts)
        {
            _cts = cts;
            Text = "Clean up images → PDF";
            Size = new Size(400, 140);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(250, 249, 248);

            _status = new Label { Text = "Reading images…", Location = new Point(20, 20), Size = new Size(350, 24) };
            _cancel = new Button
            {
                Text = "Cancel",
                Location = new Point(280, 56),
                Width = 90,
                Height = 28,
                ForeColor = Color.Black,
                BackColor = SystemColors.ButtonFace,
                UseVisualStyleBackColor = false,
            };
            _cancel.Click += (s, e) => Cancel();
            // The window's X cancels the run; only Finish() (the run ending) closes it.
            FormClosing += (s, e) =>
            {
                if (_finished) return;
                if (e.CloseReason == CloseReason.UserClosing) e.Cancel = true;
                Cancel();
            };
            // Closing the email destroys its owned windows, this one too, without a FormClosing.
            HandleDestroyed += (s, e) => { if (!_finished && !RecreatingHandle) Cancel(); };
            Controls.AddRange(new Control[] { _status, _cancel });
        }

        public void SetStatus(string text)
        {
            if (!IsDisposed && !_cts.IsCancellationRequested) _status.Text = text;
        }

        public void Finish()
        {
            _finished = true;
            if (!IsDisposed) Close();
        }

        private void Cancel()
        {
            if (_cts.IsCancellationRequested) return;
            _cancel.Enabled = false;
            _status.Text = "Cancelling…";
            _cts.Cancel();
        }
    }
}
