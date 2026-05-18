namespace OutlookAI.TaskPane.InboxReports
{
    partial class InboxReportsPane
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
            this.chatHost = new System.Windows.Forms.Panel();
            this.SuspendLayout();

            // chatHost - the controller will Dock-Fill a WebView2 into this.
            this.chatHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chatHost.Location = new System.Drawing.Point(0, 0);
            this.chatHost.Size = new System.Drawing.Size(340, 600);
            this.chatHost.TabIndex = 0;

            // InboxReportsPane
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = false;
            this.BackColor = System.Drawing.Color.FromArgb(250, 249, 248);
            this.Controls.Add(this.chatHost);
            this.Name = "InboxReportsPane";
            this.Size = new System.Drawing.Size(340, 600);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Panel chatHost;
    }
}
