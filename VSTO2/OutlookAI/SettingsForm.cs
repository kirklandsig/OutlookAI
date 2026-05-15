using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Services;

namespace OutlookAI
{
    /// <summary>
    /// Replaces the legacy nested SettingsForm. Drops API-key/model controls
    /// and adds the ChatGPT account section (Sign In / Sign Out / Refresh).
    /// Admin password gate and password change are retained.
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly CodexAuthService _auth;

        private TextBox _txtPassword;
        private Label _lblError;
        private Panel _panelSettings;
        private Label _lblAccountStatus;
        private Button _btnSignIn;
        private Button _btnSignOut;
        private Button _btnRefresh;
        private TextBox _txtNewPassword;
        private Button _btnSavePassword;

        private bool _authenticated;

        public SettingsForm()
            : this(Globals.ThisAddIn != null ? Globals.ThisAddIn.AuthService : null)
        {
        }

        public SettingsForm(CodexAuthService auth)
        {
            _auth = auth;

            Text = "OutlookAI Settings";
            Size = new Size(440, 360);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(250, 249, 248);

            BuildLoginUi();
            BuildSettingsPanel();

            if (_auth != null)
            {
                _auth.StatusChanged += OnAuthStatusChanged;
                FormClosed += (s, e) => _auth.StatusChanged -= OnAuthStatusChanged;
            }
        }

        private void BuildLoginUi()
        {
            var lblPassword = new Label
            {
                Text = "Admin Password:",
                Location = new Point(20, 20),
                AutoSize = true
            };

            _txtPassword = new TextBox
            {
                Location = new Point(20, 45),
                Width = 380,
                PasswordChar = '*'
            };

            var btnLogin = new Button
            {
                Text = "Login",
                Location = new Point(320, 75),
                Width = 80
            };
            btnLogin.Click += BtnLogin_Click;

            _lblError = new Label
            {
                Location = new Point(20, 80),
                AutoSize = true,
                ForeColor = Color.DarkRed,
                Visible = false
            };

            Controls.AddRange(new Control[] { lblPassword, _txtPassword, btnLogin, _lblError });
        }

        private void BuildSettingsPanel()
        {
            _panelSettings = new Panel
            {
                Location = new Point(0, 110),
                Size = new Size(440, 230),
                Visible = false
            };

            var grpAccount = new GroupBox
            {
                Text = "ChatGPT Account",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 10),
                Size = new Size(380, 120)
            };

            _lblAccountStatus = new Label
            {
                Location = new Point(15, 25),
                Size = new Size(350, 40),
                Font = new Font("Segoe UI", 9F),
                Text = "Status: ..."
            };

            _btnSignIn = new Button
            {
                Text = "Sign In",
                Location = new Point(15, 75),
                Width = 100
            };
            _btnSignIn.Click += async (s, e) => await SignInAsync().ConfigureAwait(false);

            _btnSignOut = new Button
            {
                Text = "Sign Out",
                Location = new Point(125, 75),
                Width = 100
            };
            _btnSignOut.Click += async (s, e) => await SignOutAsync().ConfigureAwait(false);

            _btnRefresh = new Button
            {
                Text = "Refresh",
                Location = new Point(235, 75),
                Width = 100
            };
            _btnRefresh.Click += async (s, e) => await RefreshAsync().ConfigureAwait(false);

            grpAccount.Controls.AddRange(new Control[]
            {
                _lblAccountStatus, _btnSignIn, _btnSignOut, _btnRefresh
            });

            var lblNewPassword = new Label
            {
                Text = "New Admin Password (leave blank to keep):",
                Location = new Point(20, 145),
                AutoSize = true
            };

            _txtNewPassword = new TextBox
            {
                Location = new Point(20, 165),
                Width = 240,
                PasswordChar = '*'
            };

            _btnSavePassword = new Button
            {
                Text = "Save Password",
                Location = new Point(280, 163),
                Width = 120
            };
            _btnSavePassword.Click += BtnSavePassword_Click;

            _panelSettings.Controls.AddRange(new Control[]
            {
                grpAccount, lblNewPassword, _txtNewPassword, _btnSavePassword
            });

            Controls.Add(_panelSettings);
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            if (_txtPassword.Text == Config.AdminPassword)
            {
                _authenticated = true;
                _panelSettings.Visible = true;
                _lblError.Visible = false;
                _txtPassword.Enabled = false;
                UpdateAccountUi(GetCurrentStatus());
            }
            else
            {
                _lblError.Text = "Invalid password";
                _lblError.Visible = true;
            }
        }

        private void BtnSavePassword_Click(object sender, EventArgs e)
        {
            if (!_authenticated)
            {
                return;
            }
            if (!string.IsNullOrWhiteSpace(_txtNewPassword.Text))
            {
                Config.AdminPassword = _txtNewPassword.Text;
                Config.SaveConfig();
                _txtNewPassword.Text = "";
                MessageBox.Show(
                    this,
                    "Admin password updated.",
                    "OutlookAI Settings",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private async Task SignInAsync()
        {
            if (_auth == null)
            {
                ShowAuthError("Auth service unavailable.");
                return;
            }
            SetAccountButtonsEnabled(false);
            try
            {
                await _auth.SignInAsync(CancellationToken.None).ConfigureAwait(true);
                UpdateAccountUi(_auth.GetStatus());
            }
            catch (Exception ex)
            {
                ShowAuthError(ex.Message);
            }
            finally
            {
                SetAccountButtonsEnabled(true);
            }
        }

        private async Task SignOutAsync()
        {
            if (_auth == null)
            {
                return;
            }
            var confirm = MessageBox.Show(
                this,
                "Sign OutlookAI out of ChatGPT? Other Outlook users on this server will need to sign in again.",
                "OutlookAI",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
            {
                return;
            }
            SetAccountButtonsEnabled(false);
            try
            {
                await _auth.SignOutAsync().ConfigureAwait(true);
                UpdateAccountUi(_auth.GetStatus());
            }
            catch (Exception ex)
            {
                ShowAuthError(ex.Message);
            }
            finally
            {
                SetAccountButtonsEnabled(true);
            }
        }

        private async Task RefreshAsync()
        {
            if (_auth == null)
            {
                return;
            }
            SetAccountButtonsEnabled(false);
            try
            {
                // Forcing a token fetch exercises the refresh path.
                await _auth.GetAccessTokenAsync(CancellationToken.None).ConfigureAwait(true);
                UpdateAccountUi(_auth.GetStatus());
            }
            catch (Exception ex)
            {
                ShowAuthError(ex.Message);
            }
            finally
            {
                SetAccountButtonsEnabled(true);
            }
        }

        private AuthStatus GetCurrentStatus()
        {
            return _auth != null ? _auth.GetStatus() : AuthStatus.Unauthenticated("Auth service unavailable");
        }

        private void OnAuthStatusChanged(object sender, AuthStatus status)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateAccountUi(status)));
                return;
            }
            UpdateAccountUi(status);
        }

        private void UpdateAccountUi(AuthStatus status)
        {
            switch (status.State)
            {
                case AuthState.Authenticated:
                    _lblAccountStatus.ForeColor = Color.DarkGreen;
                    _lblAccountStatus.Text = "Signed in as " + (string.IsNullOrEmpty(status.Email) ? "(unknown email)" : status.Email);
                    break;
                case AuthState.Error:
                    _lblAccountStatus.ForeColor = Color.DarkRed;
                    _lblAccountStatus.Text = "Auth error: " + status.Message;
                    break;
                default:
                    _lblAccountStatus.ForeColor = Color.DarkSlateGray;
                    _lblAccountStatus.Text = string.IsNullOrEmpty(status.Message) ? "Not signed in" : status.Message;
                    break;
            }
        }

        private void ShowAuthError(string message)
        {
            UpdateAccountUi(AuthStatus.Error(message));
        }

        private void SetAccountButtonsEnabled(bool enabled)
        {
            _btnSignIn.Enabled = enabled;
            _btnSignOut.Enabled = enabled;
            _btnRefresh.Enabled = enabled;
        }
    }
}
