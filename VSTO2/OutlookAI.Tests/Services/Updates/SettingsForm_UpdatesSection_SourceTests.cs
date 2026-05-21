using System;
using System.IO;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class SettingsForm_UpdatesSection_SourceTests
    {
        private static string FindSourceFile(params string[] parts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(parts));
                if (File.Exists(candidate)) return candidate;
                current = current.Parent;
            }
            throw new FileNotFoundException("Could not find " + Path.Combine(parts));
        }

        private static string SettingsFormSource =>
            File.ReadAllText(FindSourceFile("OutlookAI", "SettingsForm.cs"));

        [Fact]
        public void SettingsForm_HasUpdatesGroupBoxWithExpectedControls()
        {
            var src = SettingsFormSource;
            Assert.Contains("\"Updates\"", src);                       // GroupBox text
            Assert.Contains("_lblCurrentVersion", src);
            Assert.Contains("_lblLatestVersion", src);
            Assert.Contains("_lblLastChecked", src);
            Assert.Contains("_btnCheckNow", src);
            Assert.Contains("_btnInstallUpdate", src);
            Assert.Contains("_lblUpdateStatus", src);
        }

        [Fact]
        public void SettingsForm_InstallUpdateButton_DisabledUntilNewerAvailable()
        {
            var src = SettingsFormSource;
            // Button starts disabled
            Assert.Contains("_btnInstallUpdate.Enabled = false", src);
            // ... and is only enabled when the comparator says newer
            Assert.Contains("UpdateAvailability.NewerAvailable", src);
        }

        [Fact]
        public void SettingsForm_InstallClick_ShowsRdsWarningBeforeInstalling()
        {
            var src = SettingsFormSource;
            // Pin the operative copy so a future edit can't remove the
            // "all users" warning without breaking this test.
            Assert.Contains("close Outlook for ALL users", src);
            Assert.Contains("MessageBoxButtons.OKCancel", src);
        }

        [Fact]
        public void SettingsForm_UsesGitHubReleaseClient_AndPassesUserAgentFromInstalledTag()
        {
            var src = SettingsFormSource;
            Assert.Contains("new GitHubReleaseClient(", src);
            Assert.Contains("OutlookAI-Updater/", src);
        }
    }
}
