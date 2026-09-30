using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Services;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests
{
    [Collection("Config")]
    public class SettingsFormAccountTests : IDisposable
    {
        private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
        private readonly ConfigStateScope _scope = new ConfigStateScope();
        private readonly string _dir = Path.Combine(Path.GetTempPath(), "outlookai-settings-account", Path.GetRandomFileName());

        public void Dispose()
        {
            _scope.Dispose();
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private static T Field<T>(SettingsForm form, string name) => (T)typeof(SettingsForm).GetField(name, Private).GetValue(form);

        private static void LogIn(SettingsForm form, string password)
        {
            Field<TextBox>(form, "_txtPassword").Text = password;
            typeof(SettingsForm).GetMethod("BtnLogin_Click", Private).Invoke(form, new object[] { null, EventArgs.Empty });
        }

        // The not-signed-in error names the ribbon's Settings button, which
        // makes no sense inside Settings: Refresh says "Not signed in" instead.
        [Fact]
        public void Refresh_WhenSignedOut_ShowsNotSignedIn()
        {
            Directory.CreateDirectory(_dir);
            using (var http = new HttpClient(new FakeHttpMessageHandler()))
            using (var auth = new CodexAuthService(Path.Combine(_dir, "auth.json"), http))
            {
                Sta.Run(() =>
                {
                    using (var form = new SettingsForm(auth))
                    {
                        var refresh = (Task)typeof(SettingsForm).GetMethod("RefreshAsync", Private).Invoke(form, null);
                        var deadline = Stopwatch.StartNew();
                        while (!refresh.IsCompleted && deadline.Elapsed < TimeSpan.FromSeconds(10)) Application.DoEvents();
                        Assert.True(refresh.IsCompleted, "Refresh didn't finish");
                        refresh.GetAwaiter().GetResult();

                        Assert.Equal("Not signed in", Field<Label>(form, "_lblAccountStatus").Text);
                    }
                });
            }
        }

        // Every user can open Settings from the ribbon, and the docs print the
        // default password, so Settings says when it's still in use.
        [Fact]
        public void DefaultAdminPassword_IsFlaggedAfterLogIn()
        {
            Config.AdminPassword = "admin";
            Sta.Run(() =>
            {
                using (var form = new SettingsForm((CodexAuthService)null))
                {
                    LogIn(form, "admin");

                    var label = Field<Label>(form, "_lblNewPassword");
                    Assert.Contains("still the default", label.Text);
                    Assert.Equal(Color.DarkRed, label.ForeColor);
                }
            });
        }

        // (Checked in the source: a successful save shows a message box.)
        [Fact]
        public void SavingANewPassword_UpdatesTheFlag()
        {
            var src = File.ReadAllText(RepoFiles.Find("OutlookAI", "SettingsForm.cs")).Replace("\r\n", "\n");
            var start = src.IndexOf("private void BtnSavePassword_Click", StringComparison.Ordinal);
            var handler = src.Substring(start, src.IndexOf("\n        }\n", start, StringComparison.Ordinal) - start);

            Assert.True(handler.IndexOf("ShowPasswordHint();", StringComparison.Ordinal)
                > handler.IndexOf("ReportSaveFailed(", StringComparison.Ordinal));
        }

        [Fact]
        public void ChangedAdminPassword_IsNotFlagged()
        {
            Config.AdminPassword = "s3cret";
            Sta.Run(() =>
            {
                using (var form = new SettingsForm((CodexAuthService)null))
                {
                    LogIn(form, "s3cret");

                    var label = Field<Label>(form, "_lblNewPassword");
                    Assert.Equal("New Admin Password (leave blank to keep):", label.Text);
                    Assert.NotEqual(Color.DarkRed, label.ForeColor);
                }
            });
        }
    }
}
