namespace OutlookAI.TaskPane
{
    partial class AITaskPane
    {
        private System.ComponentModel.IContainer components = null;

        partial void DisposeCustomResources();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                DisposeCustomResources();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnSettings = new System.Windows.Forms.Button();
            this.grpQuickActions = new System.Windows.Forms.GroupBox();
            this.btnProofread = new System.Windows.Forms.Button();
            this.btnRevise = new System.Windows.Forms.Button();
            this.btnShorten = new System.Windows.Forms.Button();
            this.btnLengthen = new System.Windows.Forms.Button();
            this.btnFormal = new System.Windows.Forms.Button();
            this.btnFriendly = new System.Windows.Forms.Button();
            this.grpDraft = new System.Windows.Forms.GroupBox();
            this.txtDraftPrompt = new System.Windows.Forms.TextBox();
            this.btnDraft = new System.Windows.Forms.Button();
            this.btnMicDraft = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.panelResult = new System.Windows.Forms.Panel();
            this.lblResult = new System.Windows.Forms.Label();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.btnInsert = new System.Windows.Forms.Button();
            this.btnReplace = new System.Windows.Forms.Button();
            this.btnDiscard = new System.Windows.Forms.Button();
            this.grpQuickActions.SuspendLayout();
            this.grpDraft.SuspendLayout();
            this.panelResult.SuspendLayout();
            this.SuspendLayout();

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.FromArgb(0, 120, 212);
            this.lblTitle.Location = new System.Drawing.Point(10, 10);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(150, 21);
            this.lblTitle.Text = "AI Writing Assistant";

            // btnSettings
            this.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSettings.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnSettings.ForeColor = System.Drawing.Color.Black;
            this.btnSettings.BackColor = System.Drawing.Color.FromArgb(250, 249, 248);
            this.btnSettings.Location = new System.Drawing.Point(220, 8);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(30, 26);
            this.btnSettings.Text = "\u2699";
            this.btnSettings.UseVisualStyleBackColor = false;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);

            // grpQuickActions
            this.grpQuickActions.Controls.Add(this.btnProofread);
            this.grpQuickActions.Controls.Add(this.btnRevise);
            this.grpQuickActions.Controls.Add(this.btnShorten);
            this.grpQuickActions.Controls.Add(this.btnLengthen);
            this.grpQuickActions.Controls.Add(this.btnFormal);
            this.grpQuickActions.Controls.Add(this.btnFriendly);
            this.grpQuickActions.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpQuickActions.Location = new System.Drawing.Point(10, 40);
            this.grpQuickActions.Name = "grpQuickActions";
            this.grpQuickActions.Size = new System.Drawing.Size(240, 95);
            this.grpQuickActions.TabIndex = 0;
            this.grpQuickActions.TabStop = false;
            this.grpQuickActions.Text = "Quick Actions (Edit Current Email)";

            // btnProofread
            this.btnProofread.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnProofread.ForeColor = System.Drawing.Color.Black;
            this.btnProofread.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnProofread.Location = new System.Drawing.Point(10, 22);
            this.btnProofread.Size = new System.Drawing.Size(70, 28);
            this.btnProofread.Text = "Proofread";
            this.btnProofread.UseVisualStyleBackColor = false;
            this.btnProofread.Click += new System.EventHandler(this.btnProofread_Click);

            // btnRevise
            this.btnRevise.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnRevise.ForeColor = System.Drawing.Color.Black;
            this.btnRevise.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnRevise.Location = new System.Drawing.Point(85, 22);
            this.btnRevise.Size = new System.Drawing.Size(70, 28);
            this.btnRevise.Text = "Revise";
            this.btnRevise.UseVisualStyleBackColor = false;
            this.btnRevise.Click += new System.EventHandler(this.btnRevise_Click);

            // btnShorten
            this.btnShorten.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnShorten.ForeColor = System.Drawing.Color.Black;
            this.btnShorten.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnShorten.Location = new System.Drawing.Point(160, 22);
            this.btnShorten.Size = new System.Drawing.Size(70, 28);
            this.btnShorten.Text = "Shorten";
            this.btnShorten.UseVisualStyleBackColor = false;
            this.btnShorten.Click += new System.EventHandler(this.btnShorten_Click);

            // btnLengthen
            this.btnLengthen.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnLengthen.ForeColor = System.Drawing.Color.Black;
            this.btnLengthen.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnLengthen.Location = new System.Drawing.Point(10, 55);
            this.btnLengthen.Size = new System.Drawing.Size(70, 28);
            this.btnLengthen.Text = "Lengthen";
            this.btnLengthen.UseVisualStyleBackColor = false;
            this.btnLengthen.Click += new System.EventHandler(this.btnLengthen_Click);

            // btnFormal
            this.btnFormal.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnFormal.ForeColor = System.Drawing.Color.Black;
            this.btnFormal.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnFormal.Location = new System.Drawing.Point(85, 55);
            this.btnFormal.Size = new System.Drawing.Size(70, 28);
            this.btnFormal.Text = "Formal";
            this.btnFormal.UseVisualStyleBackColor = false;
            this.btnFormal.Click += new System.EventHandler(this.btnFormal_Click);

            // btnFriendly
            this.btnFriendly.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnFriendly.ForeColor = System.Drawing.Color.Black;
            this.btnFriendly.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnFriendly.Location = new System.Drawing.Point(160, 55);
            this.btnFriendly.Size = new System.Drawing.Size(70, 28);
            this.btnFriendly.Text = "Friendly";
            this.btnFriendly.UseVisualStyleBackColor = false;
            this.btnFriendly.Click += new System.EventHandler(this.btnFriendly_Click);

            // grpDraft
            this.grpDraft.Controls.Add(this.txtDraftPrompt);
            this.grpDraft.Controls.Add(this.btnMicDraft);
            this.grpDraft.Controls.Add(this.btnDraft);
            this.grpDraft.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpDraft.Location = new System.Drawing.Point(10, 140);
            this.grpDraft.Name = "grpDraft";
            this.grpDraft.Size = new System.Drawing.Size(240, 110);
            this.grpDraft.TabIndex = 1;
            this.grpDraft.TabStop = false;
            this.grpDraft.Text = "Draft New Email";

            // txtDraftPrompt
            this.txtDraftPrompt.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.txtDraftPrompt.Location = new System.Drawing.Point(10, 22);
            this.txtDraftPrompt.Multiline = true;
            this.txtDraftPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtDraftPrompt.Size = new System.Drawing.Size(185, 50);

            // btnMicDraft
            this.btnMicDraft.Font = new System.Drawing.Font("Segoe UI", 14F);
            this.btnMicDraft.ForeColor = System.Drawing.Color.Red;
            this.btnMicDraft.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnMicDraft.Location = new System.Drawing.Point(200, 22);
            this.btnMicDraft.Name = "btnMicDraft";
            this.btnMicDraft.Size = new System.Drawing.Size(30, 50);
            this.btnMicDraft.Text = "\u25CF";
            this.btnMicDraft.UseVisualStyleBackColor = false;
            this.btnMicDraft.Click += new System.EventHandler(this.btnMicDraft_Click);

            // btnDraft
            this.btnDraft.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDraft.ForeColor = System.Drawing.Color.Black;
            this.btnDraft.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDraft.Location = new System.Drawing.Point(10, 78);
            this.btnDraft.Size = new System.Drawing.Size(220, 26);
            this.btnDraft.Text = "Draft Email";
            this.btnDraft.UseVisualStyleBackColor = false;
            this.btnDraft.Click += new System.EventHandler(this.btnDraft_Click);

            // lblStatus
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblStatus.Location = new System.Drawing.Point(10, 255);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(240, 20);
            this.lblStatus.Visible = false;

            // panelResult
            this.panelResult.Controls.Add(this.lblResult);
            this.panelResult.Controls.Add(this.txtResult);
            this.panelResult.Controls.Add(this.btnInsert);
            this.panelResult.Controls.Add(this.btnReplace);
            this.panelResult.Controls.Add(this.btnDiscard);
            this.panelResult.Location = new System.Drawing.Point(10, 280);
            this.panelResult.Name = "panelResult";
            this.panelResult.Size = new System.Drawing.Size(240, 210);
            this.panelResult.Visible = false;

            // lblResult
            this.lblResult.AutoSize = true;
            this.lblResult.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblResult.Location = new System.Drawing.Point(0, 0);
            this.lblResult.Text = "Result:";

            // txtResult
            this.txtResult.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.txtResult.Location = new System.Drawing.Point(0, 20);
            this.txtResult.Multiline = true;
            this.txtResult.ReadOnly = true;
            this.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtResult.Size = new System.Drawing.Size(240, 125);

            // btnInsert
            this.btnInsert.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnInsert.ForeColor = System.Drawing.Color.Black;
            this.btnInsert.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnInsert.Location = new System.Drawing.Point(0, 150);
            this.btnInsert.Size = new System.Drawing.Size(75, 26);
            this.btnInsert.Text = "Insert";
            this.btnInsert.UseVisualStyleBackColor = false;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);

            // btnReplace
            this.btnReplace.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnReplace.ForeColor = System.Drawing.Color.Black;
            this.btnReplace.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnReplace.Location = new System.Drawing.Point(80, 150);
            this.btnReplace.Size = new System.Drawing.Size(75, 26);
            this.btnReplace.Text = "Replace";
            this.btnReplace.UseVisualStyleBackColor = false;
            this.btnReplace.Click += new System.EventHandler(this.btnReplace_Click);

            // btnDiscard
            this.btnDiscard.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnDiscard.ForeColor = System.Drawing.Color.Black;
            this.btnDiscard.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDiscard.Location = new System.Drawing.Point(160, 150);
            this.btnDiscard.Size = new System.Drawing.Size(75, 26);
            this.btnDiscard.Text = "Discard";
            this.btnDiscard.UseVisualStyleBackColor = false;
            this.btnDiscard.Click += new System.EventHandler(this.btnDiscard_Click);

            // AITaskPane
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.FromArgb(250, 249, 248);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.grpQuickActions);
            this.Controls.Add(this.grpDraft);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.panelResult);
            this.Name = "AITaskPane";
            this.Size = new System.Drawing.Size(260, 470);
            this.grpQuickActions.ResumeLayout(false);
            this.grpDraft.ResumeLayout(false);
            this.grpDraft.PerformLayout();
            this.panelResult.ResumeLayout(false);
            this.panelResult.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.GroupBox grpQuickActions;
        private System.Windows.Forms.Button btnProofread;
        private System.Windows.Forms.Button btnRevise;
        private System.Windows.Forms.Button btnShorten;
        private System.Windows.Forms.Button btnLengthen;
        private System.Windows.Forms.Button btnFormal;
        private System.Windows.Forms.Button btnFriendly;
        private System.Windows.Forms.GroupBox grpDraft;
        private System.Windows.Forms.TextBox txtDraftPrompt;
        private System.Windows.Forms.Button btnDraft;
        private System.Windows.Forms.Button btnMicDraft;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Panel panelResult;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.TextBox txtResult;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Button btnReplace;
        private System.Windows.Forms.Button btnDiscard;
    }
}
