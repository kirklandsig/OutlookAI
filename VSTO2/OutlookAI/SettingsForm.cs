using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Services;
using OutlookAI.Services.Models;
using OutlookAI.Services.Updates;

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

        // AI Behavior group (Task 28)
        private ComboBox _cmbModel;
        private Button _btnUpdateModels;
        private Label _lblModelInfo;
        private Label _lblModelCatalogStatus;
        private ComboBox _cmbReasoningEffort;
        private CheckedListBox _clbWriteTools;
        private Button _btnSaveAiSettings;
        private Label _lblAiSaved;

        // The effort the admin last chose (initially the saved one). Kept apart
        // from the dropdown so passing through a model that lacks it (Max on
        // gpt-5.5) doesn't lose it.
        private string _intendedEffort;
        private bool _effortPickedByUser;
        private bool _modelPickedByUser;
        private string _shownAccountId;

        // Updates group (Task 8)
        private GroupBox _grpUpdates;
        private Label _lblCurrentVersionCaption;
        private Label _lblCurrentVersion;
        private Label _lblLatestVersionCaption;
        private Label _lblLatestVersion;
        private Label _lblLastCheckedCaption;
        private Label _lblLastChecked;
        private Button _btnCheckNow;
        private Button _btnInstallUpdate;
        private Label _lblUpdateStatus;

        // Updater state
        private ReleaseInfo _latestRelease;
        private UpdateAvailability _availability = UpdateAvailability.NoUpdate;
        private DateTimeOffset? _lastCheckedAt;

        // Single HttpClient for the updater. Static so repeated checks reuse
        // the connection pool instead of churning sockets.
        private static readonly System.Net.Http.HttpClient _updaterHttp = new System.Net.Http.HttpClient();

        // Shared history-log handle so we don't allocate one per Append call.
        private readonly UpdateHistoryLog _history = new UpdateHistoryLog();

        private bool _authenticated;

        public SettingsForm()
            : this(Globals.ThisAddIn != null ? Globals.ThisAddIn.AuthService : null)
        {
        }

        public SettingsForm(CodexAuthService auth)
        {
            _auth = auth;

            Text = "OutlookAI Settings";
            Size = new Size(460, 620);
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
                Size = new Size(460, 500),
                Visible = false,
                AutoScroll = true
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

            BuildAiBehaviorGroup();
            BuildUpdatesGroup();

            Controls.Add(_panelSettings);
        }

        private void BuildAiBehaviorGroup()
        {
            var grpAi = new GroupBox
            {
                Text = "AI Behavior",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 200),
                Size = new Size(400, 346)
            };

            // Model dropdown + catalog refresh
            var lblModel = new Label
            {
                Text = "Model:",
                Location = new Point(15, 26),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _cmbModel = new ComboBox
            {
                Location = new Point(15, 46),
                Width = 250,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _cmbModel.SelectedIndexChanged += CmbModel_SelectedIndexChanged;

            _btnUpdateModels = new Button
            {
                Text = "Update Models",
                Location = new Point(272, 45),
                Size = new Size(113, 25),
                ForeColor = Color.Black,
                BackColor = SystemColors.ButtonFace,
                UseVisualStyleBackColor = false,
            };
            _btnUpdateModels.Click += BtnUpdateModels_Click;

            _cmbModel.SelectionChangeCommitted += (s, e) => _modelPickedByUser = true;

            // Selected model's description / retirement notice. Catalog text is
            // shown verbatim (UseMnemonic would eat '&').
            _lblModelInfo = new Label
            {
                Location = new Point(15, 74),
                Size = new Size(370, 30),
                AutoEllipsis = true,
                UseMnemonic = false,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8F, FontStyle.Regular)
            };

            // Where the list came from, or the last refresh outcome.
            _lblModelCatalogStatus = new Label
            {
                Location = new Point(15, 106),
                Size = new Size(370, 42),
                AutoEllipsis = true,
                UseMnemonic = false,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8F, FontStyle.Regular)
            };

            // Reasoning effort dropdown (re-filters when model changes)
            var lblReasoning = new Label
            {
                Text = "Reasoning effort:",
                Location = new Point(15, 152),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _cmbReasoningEffort = new ComboBox
            {
                Location = new Point(15, 172),
                Width = 370,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _cmbReasoningEffort.SelectionChangeCommitted += (s, e) =>
            {
                _effortPickedByUser = true;
                _intendedEffort = _cmbReasoningEffort.SelectedItem as string ?? _intendedEffort;
            };

            // Write-tools checklist
            var lblWriteTools = new Label
            {
                Text = "Allowed write tools:",
                Location = new Point(15, 205),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            _clbWriteTools = new CheckedListBox
            {
                Location = new Point(15, 225),
                Size = new Size(370, 80),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                IntegralHeight = false
            };
            foreach (var tool in Config.AllWriteTools)
            {
                _clbWriteTools.Items.Add(tool);
            }

            // Save button + saved indicator
            _btnSaveAiSettings = new Button
            {
                Text = "Save AI Settings",
                Location = new Point(265, 313),
                Width = 120
            };
            _btnSaveAiSettings.Click += BtnSaveAiSettings_Click;

            _lblAiSaved = new Label
            {
                Location = new Point(15, 318),
                AutoSize = true,
                ForeColor = Color.DarkGreen,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic),
                Visible = false,
                Text = "Saved."
            };

            grpAi.Controls.AddRange(new Control[]
            {
                lblModel, _cmbModel, _btnUpdateModels,
                _lblModelInfo, _lblModelCatalogStatus,
                lblReasoning, _cmbReasoningEffort,
                lblWriteTools, _clbWriteTools,
                _btnSaveAiSettings, _lblAiSaved
            });

            _panelSettings.Controls.Add(grpAi);

            // Seed initial values from current Config.
            LoadAiSettingsIntoControls();
        }

        private void LoadAiSettingsIntoControls()
        {
            _intendedEffort = Config.ReasoningEffort;
            _effortPickedByUser = false;
            _modelPickedByUser = false;

            // Show the model requests actually use. For a retired saved model
            // that is its replacement, so any Save keeps what is really in
            // effect instead of the retired slug. Selecting it also refreshes
            // the effort list and the info line.
            PopulateModelChoices(Config.EffectiveModel);
            ShowCatalogSummary();

            // Write tools
            var enabled = Config.EnabledWriteTools ?? new HashSet<string>();
            for (int i = 0; i < _clbWriteTools.Items.Count; i++)
            {
                var name = _clbWriteTools.Items[i].ToString();
                _clbWriteTools.SetItemChecked(i, enabled.Contains(name));
            }
        }

        private void CmbModel_SelectedIndexChanged(object sender, EventArgs e)
        {
            var model = _cmbModel.SelectedItem as string;
            if (string.IsNullOrEmpty(model)) return;
            // Keep the intended effort if this model takes it; otherwise show
            // Auto without forgetting it for the next model.
            RefreshReasoningEffortChoices(model, _intendedEffort);
            ShowModelInfo(model);
        }

        // Fills the model dropdown from the catalog, selecting preferModel when
        // offered, else the model requests actually use.
        private void PopulateModelChoices(string preferModel)
        {
            var catalog = Config.ModelCatalog;
            var slugs = catalog.ListedSlugs.ToList();
            // Hidden catalog models in play stay selectable (the saved one, and
            // a hidden replacement it retired to), so Save AI Settings doesn't
            // silently swap them for the first listed model.
            foreach (var inUse in new[] { Config.Model, Config.EffectiveModel })
            {
                var entry = catalog.Find(inUse);
                if (entry != null && !slugs.Contains(entry.Slug))
                {
                    slugs.Add(entry.Slug);
                }
            }

            var index = slugs.FindIndex(s => string.Equals(s, preferModel, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
            {
                index = slugs.FindIndex(s => string.Equals(s, Config.EffectiveModel, StringComparison.OrdinalIgnoreCase));
            }

            _cmbModel.BeginUpdate();
            _cmbModel.Items.Clear();
            _cmbModel.Items.AddRange(slugs.ToArray());
            _cmbModel.EndUpdate();
            // Raises SelectedIndexChanged, which refreshes efforts + info line.
            _cmbModel.SelectedIndex = index >= 0 ? index : (slugs.Count > 0 ? 0 : -1);
        }

        private void ShowModelInfo(string slug)
        {
            try
            {
                var catalog = Config.ModelCatalog;
                var entry = catalog.Find(slug);
                if (entry == null)
                {
                    _lblModelInfo.Text = "";
                    return;
                }

                var now = Config.Clock();
                var text = entry.Description.Length > 0 ? entry.DisplayName + " — " + entry.Description : entry.DisplayName;
                var warn = false;
                var upgrade = entry.Upgrade;
                if (upgrade != null && upgrade.RetirementAt.HasValue)
                {
                    warn = true;
                    var date = FormatDate(upgrade.RetirementAt.Value);
                    text = upgrade.IsRetiredAt(now)
                        ? "Retired " + date + "; requests use " + catalog.ResolveEffectiveModel(entry.Slug, now) + " instead."
                        : "Retires " + date + (string.IsNullOrEmpty(upgrade.Model) ? "" : "; switch to " + upgrade.Model) + ". " + text;
                }
                else if (upgrade != null && catalog.Find(upgrade.Model) != null)
                {
                    text += " Newer: " + upgrade.Model + ".";
                }

                // The dialog opened on the stand-in for a saved model that has
                // retired or left the list; say so until the admin picks or saves.
                var saved = Config.Model;
                if (!_modelPickedByUser
                    && !string.Equals(saved, entry.Slug, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Config.EffectiveModel, entry.Slug, StringComparison.OrdinalIgnoreCase))
                {
                    warn = true;
                    var savedEntry = catalog.Find(saved);
                    var why = savedEntry != null && savedEntry.Upgrade != null && savedEntry.Upgrade.IsRetiredAt(now)
                        ? " retired " + FormatDate(savedEntry.Upgrade.RetirementAt.Value)
                        : " is no longer offered";
                    text = "Saved model " + saved + why + "; requests use " + entry.Slug
                        + ". Click Save AI Settings to keep it. " + text;
                }

                _lblModelInfo.ForeColor = warn ? Color.DarkRed : Color.DimGray;
                _lblModelInfo.Text = text;
            }
            catch (Exception ex)
            {
                // A bad cache must never keep Settings (and Update Models) from opening.
                _lblModelInfo.Text = "";
                OutlookAI.Diagnostics.TraceLog.Write("ShowModelInfo failed: " + ex.Message, "Settings");
            }
        }

        private void ShowCatalogSummary()
        {
            try
            {
                var catalog = Config.ModelCatalog;
                var color = Color.DimGray;
                var text = catalog.IsBuiltIn || !catalog.FetchedAt.HasValue
                    // AsOf is a capture date (midnight UTC), not a moment: show it as-is.
                    ? "Model list: built-in (as of "
                        + BuiltInModelCatalog.AsOf.UtcDateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        + "). Update Models fetches the current list from ChatGPT."
                    : "Model list: " + catalog.ListedSlugs.Count + " models from ChatGPT, updated "
                        + FormatDateTime(catalog.FetchedAt.Value) + ".";

                _shownAccountId = _auth != null ? _auth.GetStatus().AccountId : "";
                if (!catalog.IsBuiltIn && !string.IsNullOrEmpty(catalog.AccountId) && !string.IsNullOrEmpty(_shownAccountId)
                    && !string.Equals(catalog.AccountId, _shownAccountId, StringComparison.Ordinal))
                {
                    color = Color.DarkOrange;
                    text += " It was fetched for a different ChatGPT account; click Update Models.";
                }
                ShowCatalogStatus(text, color);
            }
            catch (Exception ex)
            {
                _lblModelCatalogStatus.Text = "";
                OutlookAI.Diagnostics.TraceLog.Write("ShowCatalogSummary failed: " + ex.Message, "Settings");
            }
        }

        private void ShowCatalogStatus(string text, Color color)
        {
            _lblModelCatalogStatus.ForeColor = color;
            _lblModelCatalogStatus.Text = text;
        }

        // Invariant culture: the Gregorian calendar can format any date the
        // catalog parser accepts; e.g. ar-SA's UmAlQura throws past 2077.
        private static string FormatDate(DateTimeOffset value)
        {
            return value.ToLocalTime().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static string FormatDateTime(DateTimeOffset value)
        {
            return value.ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        private async void BtnUpdateModels_Click(object sender, EventArgs e)
        {
            if (!_authenticated) return;

            _btnUpdateModels.Enabled = false;
            ShowCatalogStatus("Updating the model list from ChatGPT…", Color.DimGray);
            try
            {
                using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(90)))
                {
                    var updater = new ModelCatalogUpdater(_updaterHttp, new ModelCatalogStore());
                    // Models named in any config layer (plus what's in use) stay
                    // valid even if this account's list happens to omit them.
                    var keep = Config.ModelsNamedInConfig.Concat(new[] { Config.Model, Config.EffectiveModel });
                    var result = await updater.UpdateAsync(
                        _auth, Config.ModelCatalog, Config.ModelCatalogClientVersion, keep, cts.Token);
                    if (!result.Succeeded)
                    {
                        ShowCatalogStatus("Update failed: " + result.Error + " The current list is unchanged.", Color.DarkRed);
                        return;
                    }

                    // Swap first so the list takes effect even if the dialog is
                    // closing. Then re-apply config.xml against it (a model or
                    // effort the old list rejected may be valid now) and let open
                    // panes refresh their effort dropdowns.
                    Config.ModelCatalog = result.Catalog;
                    Config.ReloadConfigFiles();
                    Config.NotifyAiSettingsChanged();

                    if (!_effortPickedByUser) _intendedEffort = Config.ReasoningEffort;
                    PopulateModelChoices(_modelPickedByUser ? _cmbModel.SelectedItem as string : Config.EffectiveModel);

                    var warning = DescribeSaveProblem(result.Save);
                    ShowCatalogStatus(
                        DescribeCatalogUpdate(result) + (warning == null ? "" : " See the warning."),
                        warning == null ? Color.DarkGreen : Color.DarkOrange);
                    if (warning != null)
                    {
                        MessageBox.Show(this, warning, "Update Models", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                // Settings form was closed mid-await; nothing more to do.
            }
            catch (OperationCanceledException)
            {
                try { ShowCatalogStatus("Update timed out. The current list is unchanged.", Color.DarkRed); } catch { }
            }
            catch (Exception ex)
            {
                try { ShowCatalogStatus("Update failed: " + ex.Message, Color.DarkRed); } catch { }
            }
            finally
            {
                try { _btnUpdateModels.Enabled = true; } catch { }
            }
        }

        private static string DescribeCatalogUpdate(ModelCatalogUpdateResult result)
        {
            var changes = new List<string>();
            if (result.Added.Count > 0) changes.Add("new: " + string.Join(", ", result.Added));
            if (result.Removed.Count > 0) changes.Add("removed: " + string.Join(", ", result.Removed));
            var text = result.SupersededByNewer
                ? "Another Outlook session saved a newer list moments ago; using it (" + result.Catalog.ListedSlugs.Count + " models)."
                : "Updated: " + result.Catalog.ListedSlugs.Count + " models ("
                    + (changes.Count > 0 ? string.Join("; ", changes) : "no changes") + ").";
            if (result.KeptUnlisted.Count > 0)
            {
                text += " Not in ChatGPT's list for this account, still used as configured: "
                    + string.Join(", ", result.KeptUnlisted) + ".";
            }
            if (!result.ServerEffortsProbed && !result.SupersededByNewer)
            {
                text += " Couldn't confirm effort levels; kept the last known set.";
            }
            return text;
        }

        // Null when both copies were written.
        private static string DescribeSaveProblem(ModelCatalogSaveResult save)
        {
            if (!save.SharedSaved && !save.UserSaved)
            {
                return "The new model list could not be saved, so it only applies until Outlook restarts.\n\n"
                    + "Shared copy: " + save.SharedError + "\n"
                    + "Your copy: " + save.UserError;
            }
            if (!save.SharedSaved)
            {
                return "The new model list was saved for your Windows account only. Other users on this machine "
                    + "keep the previous list until the shared copy can be written:\n\n" + save.SharedError;
            }
            return null;
        }

        private void RefreshReasoningEffortChoices(string model, string preferEffort)
        {
            var efforts = Config.ReasoningEffortsForModel(model);
            _cmbReasoningEffort.BeginUpdate();
            _cmbReasoningEffort.Items.Clear();
            _cmbReasoningEffort.Items.AddRange(efforts);
            // Snap to preferred if valid, else first entry.
            var idx = Array.IndexOf(efforts, preferEffort);
            _cmbReasoningEffort.SelectedIndex = idx >= 0 ? idx : 0;
            _cmbReasoningEffort.EndUpdate();
        }

        private void BtnSaveAiSettings_Click(object sender, EventArgs e)
        {
            if (!_authenticated) return;

            var pickedModel = _cmbModel.SelectedItem as string;
            if (!string.IsNullOrEmpty(pickedModel))
            {
                Config.Model = pickedModel;
            }

            var pickedEffort = _cmbReasoningEffort.SelectedItem as string;
            if (!string.IsNullOrEmpty(pickedEffort))
            {
                Config.ReasoningEffort = pickedEffort;
            }

            // EnabledWriteTools - collect every checked entry.
            var newSet = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < _clbWriteTools.Items.Count; i++)
            {
                if (_clbWriteTools.GetItemChecked(i))
                {
                    newSet.Add(_clbWriteTools.Items[i].ToString());
                }
            }
            Config.EnabledWriteTools = newSet;
            // Master switch is derived: any write tool checked = true.
            Config.WriteToolsEnabled = newSet.Count > 0;

            Config.SaveConfig();
            // Open panes re-read their effort lists for the (maybe new) model.
            Config.NotifyAiSettingsChanged();

            // What's shown is now what's saved: drop any "saved model retired"
            // notice and carry the saved effort through later model switches.
            _intendedEffort = Config.ReasoningEffort;
            _effortPickedByUser = false;
            _modelPickedByUser = false;
            if (!string.IsNullOrEmpty(pickedModel)) ShowModelInfo(pickedModel);

            _lblAiSaved.Visible = true;
            // Auto-hide the saved indicator after a short delay so repeated
            // saves still give clear visual feedback.
            var t = new System.Windows.Forms.Timer { Interval = 2500 };
            t.Tick += (s2, e2) => { _lblAiSaved.Visible = false; t.Stop(); t.Dispose(); };
            t.Start();
        }

        private void BuildUpdatesGroup()
        {
            // AI Behavior group lives at (20, 200) with height 346, so its
            // bottom edge is y=546. Place Updates 12 px below that for visual
            // separation; the panel scrolls if the form is short.
            _grpUpdates = new GroupBox
            {
                Text = "Updates",
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Location = new Point(20, 558),
                Size = new Size(400, 180),
                ForeColor = Color.Black,
            };

            _lblCurrentVersionCaption = new Label { Text = "Current:",      Location = new Point(12, 24), AutoSize = true, ForeColor = Color.Black };
            _lblCurrentVersion        = new Label { Text = "—",             Location = new Point(96, 24), AutoSize = true, ForeColor = Color.Black };

            _lblLatestVersionCaption  = new Label { Text = "Latest:",       Location = new Point(12, 48), AutoSize = true, ForeColor = Color.Black };
            _lblLatestVersion         = new Label { Text = "—",             Location = new Point(96, 48), AutoSize = true, ForeColor = Color.Black };

            _lblLastCheckedCaption    = new Label { Text = "Last checked:", Location = new Point(12, 72), AutoSize = true, ForeColor = Color.Black };
            _lblLastChecked           = new Label { Text = "—",             Location = new Point(96, 72), AutoSize = true, ForeColor = Color.Black };

            _btnCheckNow = new Button
            {
                Text = "Check Now",
                Location = new Point(12, 100),
                Size = new Size(110, 28),
                ForeColor = Color.Black,
                BackColor = SystemColors.ButtonFace,
                UseVisualStyleBackColor = false,
            };
            _btnCheckNow.Click += BtnCheckNow_Click;

            _btnInstallUpdate = new Button
            {
                Text = "Install Update",
                Location = new Point(130, 100),
                Size = new Size(130, 28),
                ForeColor = Color.Black,
                BackColor = SystemColors.ButtonFace,
                UseVisualStyleBackColor = false,
                Enabled = false,
            };
            _btnInstallUpdate.Click += BtnInstallUpdate_Click;

            _lblUpdateStatus = new Label
            {
                Text = "",
                Location = new Point(12, 138),
                AutoSize = false,
                Size = new Size(375, 32),
                ForeColor = Color.Black,
            };

            _grpUpdates.Controls.AddRange(new Control[]
            {
                _lblCurrentVersionCaption, _lblCurrentVersion,
                _lblLatestVersionCaption,  _lblLatestVersion,
                _lblLastCheckedCaption,    _lblLastChecked,
                _btnCheckNow, _btnInstallUpdate, _lblUpdateStatus,
            });
            _panelSettings.Controls.Add(_grpUpdates);

            // Populate current version from disk so the user sees it as soon
            // as the form opens, before any "Check Now" click.
            _lblCurrentVersion.Text = UpdateManifest.LoadFromInstallDir().Tag;
        }

        private async void BtnCheckNow_Click(object sender, EventArgs e)
        {
            _btnCheckNow.Enabled = false;
            _btnInstallUpdate.Enabled = false;
            _lblUpdateStatus.Text = "Checking…";

            try
            {
                var installed = UpdateManifest.LoadFromInstallDir();
                var ua = "OutlookAI-Updater/" + (installed.IsDevBuild ? "dev" : installed.Tag);
                var client = new GitHubReleaseClient(_updaterHttp, "kirklandsig/OutlookAI", ua);

                var result = await client.GetLatestStableAsync(CancellationToken.None);
                _lastCheckedAt = DateTimeOffset.Now;
                _lblLastChecked.Text = _lastCheckedAt.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm");

                switch (result)
                {
                    case ReleaseFound found:
                        _latestRelease = found.Info;
                        _lblLatestVersion.Text = found.Info.Tag;
                        _availability = VersionComparator.Compare(installed.Tag, found.Info.Tag);
                        _btnInstallUpdate.Enabled = _availability == UpdateAvailability.NewerAvailable;
                        _lblUpdateStatus.Text = _availability == UpdateAvailability.NewerAvailable
                            ? ("Update available: " + found.Info.Tag)
                            : (_availability == UpdateAvailability.NoUpdate ? "Already up to date." :
                               _availability == UpdateAvailability.OlderThanInstalled ? "Latest is older than installed (unusual)." :
                               "Latest tag could not be compared to installed version.");
                        break;
                    case NoReleasesAvailable _:
                        _latestRelease = null;
                        _availability = UpdateAvailability.NoReleases;
                        _lblLatestVersion.Text = "—";
                        _lblUpdateStatus.Text = "No releases published yet on GitHub.";
                        break;
                    case RateLimited rl:
                        _latestRelease = null;
                        _availability = UpdateAvailability.NotComparable;
                        _lblLatestVersion.Text = "—";
                        _lblUpdateStatus.Text = "GitHub rate limit hit. Try again after " + rl.ResetAt.ToLocalTime().ToString("HH:mm") + ".";
                        break;
                    case NetworkError ne:
                        _latestRelease = null;
                        _availability = UpdateAvailability.NotComparable;
                        _lblLatestVersion.Text = "—";
                        _lblUpdateStatus.Text = "Could not reach GitHub: " + ne.Detail;
                        break;
                }

                try
                {
                    _history.Append("check",
                        result.GetType().Name.ToLowerInvariant(),
                        (_latestRelease != null ? _latestRelease.Tag : ""),
                        _lblUpdateStatus.Text);
                }
                catch { /* Logging is best-effort; never break the update flow. */ }
            }
            catch (ObjectDisposedException)
            {
                // Settings form was closed mid-await; nothing more to do.
            }
            catch (Exception ex)
            {
                try { _lblUpdateStatus.Text = "Update check failed: " + ex.Message; } catch { }
                try { _history.Append("check", "exception", "", ex.Message); } catch { }
            }
            finally
            {
                try { _btnCheckNow.Enabled = true; } catch { }
            }
        }

        private async void BtnInstallUpdate_Click(object sender, EventArgs e)
        {
            if (_latestRelease == null || _availability != UpdateAvailability.NewerAvailable) return;

            var confirm = MessageBox.Show(
                text:
                    "Install OutlookAI " + _latestRelease.Tag + ".\n\n" +
                    "This will:\n" +
                    "  • close Outlook for ALL users currently on this server\n" +
                    "  • run the OutlookAI installer with administrator privileges\n" +
                    "  • leave Outlook closed when finished — everyone reopens manually\n\n" +
                    "Have you given users a heads-up?",
                caption: "Install Update",
                buttons: MessageBoxButtons.OKCancel,
                icon: MessageBoxIcon.Warning,
                defaultButton: MessageBoxDefaultButton.Button2);
            if (confirm != DialogResult.OK) return;

            _btnInstallUpdate.Enabled = false;
            _btnCheckNow.Enabled = false;
            _lblUpdateStatus.Text = "Downloading…";

            var launchedSuccessfully = false;
            try
            {
                var downloader = new UpdateDownloader(_updaterHttp);
                var dl = await downloader.DownloadAsync(_latestRelease, null, CancellationToken.None);

                if (!(dl is DownloadSuccess success))
                {
                    // C# 7.3: classic switch (switch expressions require C# 8).
                    switch (dl)
                    {
                        case HashMismatch _:
                            _lblUpdateStatus.Text = "Downloaded file failed integrity check. Aborting.";
                            break;
                        case MissingInstallerScript _:
                            _lblUpdateStatus.Text = "Update package is malformed (no installer). Please file a bug.";
                            break;
                        case DownloadFailed df:
                            _lblUpdateStatus.Text = "Download failed: " + df.Detail;
                            break;
                        case Cancelled _:
                            _lblUpdateStatus.Text = "Cancelled.";
                            break;
                        default:
                            _lblUpdateStatus.Text = "Unknown download result.";
                            break;
                    }
                    try { _history.Append("download", "failed", _latestRelease.Tag, _lblUpdateStatus.Text); } catch { }
                    return;
                }
                try { _history.Append("download", "ok", _latestRelease.Tag, "sha256_ok"); } catch { }

                // Write sentinel; cleared by ThisAddIn.Startup on next Outlook start.
                try
                {
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(UpdatePaths.InProgressSentinel));
                    System.IO.File.WriteAllText(UpdatePaths.InProgressSentinel, _latestRelease.Tag);
                }
                catch { }

                var installer = new UpdateInstaller();
                var launch = installer.LaunchElevatedInstall(success);

                switch (launch)
                {
                    case Launched l:
                        _lblUpdateStatus.Text = "Installer launched (PID " + l.Pid + "). Outlook will close shortly to apply the update.";
                        try { _history.Append("launch", "launched", _latestRelease.Tag, "pid=" + l.Pid); } catch { }
                        launchedSuccessfully = true;
                        break;
                    case UacDeclined _:
                        _lblUpdateStatus.Text = "Update cancelled — administrator privileges required.";
                        try { System.IO.File.Delete(UpdatePaths.InProgressSentinel); } catch { }
                        try { _history.Append("launch", "uac_declined", _latestRelease.Tag, ""); } catch { }
                        break;
                    case LaunchFailed lf:
                        _lblUpdateStatus.Text = "Failed to launch installer: " + lf.Detail;
                        try { System.IO.File.Delete(UpdatePaths.InProgressSentinel); } catch { }
                        try { _history.Append("launch", "failed", _latestRelease.Tag, lf.Detail); } catch { }
                        break;
                }
            }
            catch (ObjectDisposedException)
            {
                // Settings form was closed mid-await; nothing more to do.
            }
            catch (Exception ex)
            {
                try { _lblUpdateStatus.Text = "Install failed: " + ex.Message; } catch { }
                try { _history.Append("install", "exception", _latestRelease != null ? _latestRelease.Tag : "", ex.Message); } catch { }
                // Sentinel may have been written; try to clean it up so we don't lie to the reconciler.
                try { System.IO.File.Delete(UpdatePaths.InProgressSentinel); } catch { }
            }
            finally
            {
                if (!launchedSuccessfully)
                {
                    try { _btnInstallUpdate.Enabled = true; } catch { }
                    try { _btnCheckNow.Enabled = true; } catch { }
                }
            }
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

            // Another ChatGPT account may be offered a different model list.
            if (status.State == AuthState.Authenticated && _lblModelCatalogStatus != null
                && !string.Equals(status.AccountId, _shownAccountId, StringComparison.Ordinal))
            {
                ShowCatalogSummary();
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
