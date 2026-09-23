using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using OutlookAI.Services;
using OutlookAI.Services.Models;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Models
{
    [Collection("Config")]
    public class SettingsFormModelCatalogTests : IDisposable
    {
        private readonly ConfigStateScope _scope = new ConfigStateScope();

        public void Dispose()
        {
            _scope.Dispose();
        }

        private static T Field<T>(object owner, string name)
        {
            return (T)owner.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(owner);
        }

        // What an admin's pick records (the dropdowns' SelectionChangeCommitted,
        // which setting SelectedItem in code doesn't raise).
        private static void Pick(SettingsForm form, string combo, string value)
        {
            var flags = BindingFlags.NonPublic | BindingFlags.Instance;
            Field<ComboBox>(form, combo).SelectedItem = value;
            if (combo == "_cmbModel")
            {
                typeof(SettingsForm).GetField("_modelPickedByUser", flags).SetValue(form, true);
            }
            else
            {
                typeof(SettingsForm).GetField("_intendedEffort", flags).SetValue(form, value);
            }
        }

        // Opens Settings signed in, runs the admin's steps and returns what each
        // Save AI Settings stored (recorded instead of written to disk).
        private static List<string[]> Saves(Action<SettingsForm> steps)
        {
            var saves = new List<string[]>();
            var save = SettingsForm.SaveSettings;
            SettingsForm.SaveSettings = (changed, toolsBefore) => { saves.Add(changed); return null; };
            try
            {
                Sta.Run(() =>
                {
                    using (var form = new SettingsForm((CodexAuthService)null))
                    {
                        typeof(SettingsForm).GetField("_authenticated", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, true);
                        steps(form);
                    }
                });
            }
            finally
            {
                SettingsForm.SaveSettings = save;
            }
            return saves;
        }

        private static void Save(SettingsForm form)
        {
            typeof(SettingsForm).GetMethod("BtnSaveAiSettings_Click", BindingFlags.NonPublic | BindingFlags.Instance)
                .Invoke(form, new object[] { null, EventArgs.Empty });
        }

        // What Update Models does once it has reloaded the settings.
        private static void Reload(SettingsForm form)
        {
            typeof(SettingsForm).GetMethod("ShowReloadedSettings", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, null);
        }

        private static object Shown(SettingsForm form, string combo)
        {
            return Field<ComboBox>(form, combo).SelectedItem;
        }

        private static CheckedListBox Tools(SettingsForm form)
        {
            return Field<CheckedListBox>(form, "_clbWriteTools");
        }

        [Fact]
        public void RetiredSavedModel_ShowsTheModelInUse_AndKeepsTheSavedEffort()
        {
            // After 2026-10-14 requests go to gpt-5.6-sol at Max. Settings must
            // show that, not gpt-5.5 with Max snapped to None (which any Save
            // would then persist for every user).
            Config.Model = "gpt-5.5";
            Config.ReasoningEffort = "Max";
            Config.Clock = () => new DateTimeOffset(2026, 10, 15, 0, 0, 0, TimeSpan.Zero);

            Sta.Run(() =>
            {
                using (var form = new SettingsForm((CodexAuthService)null))
                {
                    Assert.Equal("gpt-5.6-sol", Field<ComboBox>(form, "_cmbModel").SelectedItem);
                    Assert.Equal("Max", Field<ComboBox>(form, "_cmbReasoningEffort").SelectedItem);
                    var info = Field<Label>(form, "_lblModelInfo").Text;
                    Assert.Contains("gpt-5.5", info);
                    Assert.Contains("retired", info);
                }
            });
        }

        [Fact]
        public void HiddenRetirementReplacement_IsSelectable_AndSelected()
        {
            // Otherwise the picker falls back to the first listed model and an
            // unrelated Save persists that instead of the working route.
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("top", 1, new[] { "low" }),
                TestCatalogs.Entry("a", 2, new[] { "low" },
                    upgrade: new ModelUpgradeInfo("b", TestCatalogs.FixedNow.AddDays(-1))),
                TestCatalogs.Entry("b", 3, new[] { "low" }, listed: false));
            Config.Model = "a";

            Sta.Run(() =>
            {
                using (var form = new SettingsForm((CodexAuthService)null))
                {
                    Assert.Equal("b", Field<ComboBox>(form, "_cmbModel").SelectedItem);
                    Assert.Contains("Saved model a retired", Field<Label>(form, "_lblModelInfo").Text);
                }
            });
        }

        [Fact]
        public void SwitchingThroughAModelWithoutTheEffort_RestoresItAfterwards()
        {
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "Max";

            Sta.Run(() =>
            {
                using (var form = new SettingsForm((CodexAuthService)null))
                {
                    var model = Field<ComboBox>(form, "_cmbModel");
                    var effort = Field<ComboBox>(form, "_cmbReasoningEffort");
                    Assert.Equal("Max", effort.SelectedItem);

                    model.SelectedItem = "gpt-5.5";
                    Assert.Equal("Auto", effort.SelectedItem);

                    model.SelectedItem = "gpt-6-sol";
                    Assert.Equal("Max", effort.SelectedItem);
                }
            });
        }

        [Fact]
        public void Opens_UnderANonGregorianCulture_WithFarFutureDates()
        {
            // UmAlQura can't represent 2099; culture-sensitive formatting threw
            // inside the constructor and Settings never opened. (Mid-year
            // timestamps: the dialog shows local dates.)
            Config.ModelCatalog = new ModelCatalog(
                new[]
                {
                    TestCatalogs.Entry("m", 1, new[] { "low" },
                        upgrade: new ModelUpgradeInfo("n", new DateTimeOffset(2150, 6, 15, 12, 0, 0, TimeSpan.Zero))),
                    TestCatalogs.Entry("n", 2, new[] { "low" }),
                },
                null, ModelCatalog.SourceChatGpt, new DateTimeOffset(2099, 6, 15, 12, 0, 0, TimeSpan.Zero), "0.155.1");
            Config.Model = "m";

            Sta.Run(() =>
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("ar-SA");
                using (var form = new SettingsForm((CodexAuthService)null))
                {
                    Assert.Contains("2099", Field<Label>(form, "_lblModelCatalogStatus").Text);
                    Assert.Contains("2150", Field<Label>(form, "_lblModelInfo").Text);
                }
            });
        }

        [Fact]
        public void SaveAiSettings_SavesOnlyWhatWasChangedInTheDialog()
        {
            // Another admin may have saved other settings for every user since
            // this dialog opened; saving them all again would undo that.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";

            var saves = Saves(form =>
            {
                // Re-picking what's saved, or unchecking and rechecking a tool, isn't a change.
                Pick(form, "_cmbModel", "gpt-6-astra");
                Pick(form, "_cmbReasoningEffort", "High");
                Tools(form).SetItemChecked(1, false);
                Tools(form).SetItemChecked(1, true);
                Save(form);
                Pick(form, "_cmbReasoningEffort", "Low");
                Save(form);
                Tools(form).SetItemChecked(0, false);
                Save(form);
            });

            Assert.Empty(saves[0]);
            Assert.Equal(new[] { "ReasoningEffort" }, saves[1]);
            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[2]);
        }

        [Fact]
        public void WriteToolsSwitchedOff_ShowNoneChecked_AndAnUnrelatedSaveLeavesThemOff()
        {
            // An older save may have left a tool list behind the off switch;
            // showing it checked would make the next Save switch it back on.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";
            Config.WriteToolsEnabled = false;

            var saves = Saves(form =>
            {
                Assert.Empty(Tools(form).CheckedItems);
                Pick(form, "_cmbReasoningEffort", "Low");
                Save(form);
            });

            Assert.Equal(new[] { "ReasoningEffort" }, saves[0]);
            Assert.False(Config.WriteToolsEnabled);
        }

        [Fact]
        public void AfterAReload_UntouchedControlsFollowIt_AndTheAdminsEditsStay()
        {
            // Update Models reloaded another admin's change (write tools off);
            // saving the admin's effort edit must not undo it.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";

            var saves = Saves(form =>
            {
                Pick(form, "_cmbReasoningEffort", "Low");

                Config.WriteToolsEnabled = false;   // reloaded: another admin's change
                Config.EnabledWriteTools = new HashSet<string>();
                Reload(form);

                Assert.Empty(Tools(form).CheckedItems);
                Assert.Equal("Low", Shown(form, "_cmbReasoningEffort"));
                Save(form);
            });

            Assert.Equal(new[] { "ReasoningEffort" }, saves[0]);
        }

        [Fact]
        public void AfterAReload_TheAdminsModelAndToolEditsStay()
        {
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";

            var saves = Saves(form =>
            {
                Pick(form, "_cmbModel", "gpt-6-sol");
                Tools(form).SetItemChecked(0, false);   // the admin unchecks a tool

                Config.Model = "gpt-6-luna";   // reloaded: another admin's change
                Reload(form);

                Assert.Equal("gpt-6-sol", Shown(form, "_cmbModel"));
                Assert.False(Tools(form).GetItemChecked(0));
                Save(form);
            });

            Assert.Equal(new[] { "Model", "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
            Assert.Equal("gpt-6-sol", Config.Model);
        }

        [Theory]
        [InlineData("model")]
        [InlineData("effort")]
        [InlineData("tools")]
        public void AnEditTheAdminUndid_FollowsAReload_SoAnUnrelatedSaveKeepsAnotherAdminsChange(string setting)
        {
            // The admin changes a setting and changes it back; another admin saves
            // a new value, and Update Models reloads it. Saving something else
            // must not write the old value back.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";

            var saves = Saves(form =>
            {
                var tools = Tools(form);
                switch (setting)
                {
                    case "model":
                        Pick(form, "_cmbModel", "gpt-6-sol");
                        Pick(form, "_cmbModel", "gpt-6-astra");
                        Config.Model = "gpt-6-sol";
                        break;
                    case "effort":
                        Pick(form, "_cmbReasoningEffort", "Low");
                        Pick(form, "_cmbReasoningEffort", "High");
                        Config.ReasoningEffort = "Max";
                        break;
                    default:
                        tools.SetItemChecked(0, false);
                        tools.SetItemChecked(0, true);
                        Config.WriteToolsEnabled = false;
                        Config.EnabledWriteTools = new HashSet<string>();
                        break;
                }
                Reload(form);

                Assert.Equal(Config.Model, Shown(form, "_cmbModel"));
                Assert.Equal(Config.ReasoningEffort, Shown(form, "_cmbReasoningEffort"));
                Assert.Equal(Config.WriteToolsEnabled ? tools.Items.Count : 0, tools.CheckedItems.Count);
                if (setting == "tools") Pick(form, "_cmbReasoningEffort", "Low");
                else tools.SetItemChecked(0, false);
                Save(form);
            });

            Assert.Equal(setting == "tools" ? new[] { "ReasoningEffort" } : new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
        }

        [Theory]
        [InlineData("model")]
        [InlineData("effort")]
        [InlineData("tools")]
        public void AfterAReload_AnEditCountsAgainstWhatWasReloaded(string setting)
        {
            // The admin changes a setting and another admin saves the same value:
            // once that is reloaded, changing it back is a change to save.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";

            var saves = Saves(form =>
            {
                var tools = Tools(form);
                switch (setting)
                {
                    case "model":
                        Pick(form, "_cmbModel", "gpt-6-sol");
                        Config.Model = "gpt-6-sol";
                        Reload(form);
                        Pick(form, "_cmbModel", "gpt-6-astra");
                        break;
                    case "effort":
                        Pick(form, "_cmbReasoningEffort", "Low");
                        Config.ReasoningEffort = "Low";
                        Reload(form);
                        Pick(form, "_cmbReasoningEffort", "High");
                        break;
                    default:
                        tools.SetItemChecked(0, false);
                        var others = new HashSet<string>(Config.EnabledWriteTools);
                        others.Remove(tools.Items[0].ToString());
                        Config.EnabledWriteTools = others;
                        Reload(form);
                        tools.SetItemChecked(0, true);
                        break;
                }
                Save(form);
            });

            Assert.Equal(
                setting == "model" ? new[] { "Model" }
                    : setting == "effort" ? new[] { "ReasoningEffort" }
                    : new[] { "EnabledWriteTools", "WriteToolsEnabled" },
                saves[0]);
            Assert.Equal("gpt-6-astra", Config.Model);
            Assert.Equal("High", Config.ReasoningEffort);
            Assert.Equal(Config.AllWriteTools.Length, Config.EnabledWriteTools.Count);
        }

        [Theory]
        [InlineData("model")]
        [InlineData("effort")]
        public void AfterAReload_WhatTheDialogShowsOnItsOwn_IsNotSaved(string setting)
        {
            // The admin's pick was undone; the reloaded settings name a retired
            // model (shown as its replacement) or an effort the model lacks
            // (shown as Auto). Neither is the admin's change to save.
            Config.ResetDefaults();
            Config.Model = setting == "model" ? "gpt-6-astra" : "gpt-5.5";
            Config.ReasoningEffort = "High";
            if (setting == "model") Config.Clock = () => new DateTimeOffset(2026, 10, 15, 0, 0, 0, TimeSpan.Zero);

            var saves = Saves(form =>
            {
                if (setting == "model")
                {
                    Pick(form, "_cmbModel", "gpt-6-sol");
                    Pick(form, "_cmbModel", "gpt-6-astra");
                    Config.Model = "gpt-5.5";
                }
                else
                {
                    Pick(form, "_cmbReasoningEffort", "Low");
                    Pick(form, "_cmbReasoningEffort", "High");
                    Config.ReasoningEffort = "Max";
                }
                Reload(form);

                Assert.Equal(setting == "model" ? "gpt-5.6-sol" : "gpt-5.5", Shown(form, "_cmbModel"));
                Assert.Equal(setting == "model" ? "High" : "Auto", Shown(form, "_cmbReasoningEffort"));
                Tools(form).SetItemChecked(0, false);
                Save(form);
            });

            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal(setting == "model" ? "High" : "Max", Config.ReasoningEffort);
        }

        [Fact]
        public void APickTheUpdatedModelListDropped_IsNotSaved()
        {
            // The dialog falls back to the model in use; that isn't the admin's pick.
            Config.ResetDefaults();
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("a", 1, new[] { "low" },
                    upgrade: new ModelUpgradeInfo("b", TestCatalogs.FixedNow.AddDays(-1))),
                TestCatalogs.Entry("b", 2, new[] { "low" }),
                TestCatalogs.Entry("c", 3, new[] { "low" }));
            Config.Model = "a";

            var saves = Saves(form =>
            {
                Pick(form, "_cmbModel", "c");
                Config.ModelCatalog = TestCatalogs.Catalog(
                    TestCatalogs.Entry("a", 1, new[] { "low" },
                        upgrade: new ModelUpgradeInfo("b", TestCatalogs.FixedNow.AddDays(-1))),
                    TestCatalogs.Entry("b", 2, new[] { "low" }));
                Reload(form);

                Assert.Equal("b", Shown(form, "_cmbModel"));
                Assert.Contains("Saved model a retired", Field<Label>(form, "_lblModelInfo").Text);
                Tools(form).SetItemChecked(0, false);
                Save(form);
            });

            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
            Assert.Equal("a", Config.Model);
        }

        [Fact]
        public void AnEffortTheUpdatedModelListDropped_IsNotSaved()
        {
            Config.ResetDefaults();
            Config.ModelCatalog = TestCatalogs.Catalog(TestCatalogs.Entry("a", 1, new[] { "low", "high" }));
            Config.Model = "a";

            var saves = Saves(form =>
            {
                Pick(form, "_cmbReasoningEffort", "High");
                Config.ModelCatalog = TestCatalogs.Catalog(TestCatalogs.Entry("a", 1, new[] { "low" }));
                Reload(form);

                Assert.Equal("Auto", Shown(form, "_cmbReasoningEffort"));
                Tools(form).SetItemChecked(0, false);
                Save(form);
            });

            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
            Assert.Equal("Auto", Config.ReasoningEffort);
        }

        [Fact]
        public void AfterASave_ChangingASettingBack_IsAChange()
        {
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";

            var saves = Saves(form =>
            {
                Pick(form, "_cmbModel", "gpt-6-sol");
                Save(form);
                Pick(form, "_cmbModel", "gpt-6-astra");
                Save(form);
                Pick(form, "_cmbReasoningEffort", "Low");
                Save(form);
                Pick(form, "_cmbReasoningEffort", "High");
                Save(form);
                Tools(form).SetItemChecked(0, false);
                Save(form);
                Tools(form).SetItemChecked(0, true);
                Save(form);
                Save(form);
            });

            Assert.Equal(new[] { "Model" }, saves[0]);
            Assert.Equal(new[] { "Model" }, saves[1]);
            Assert.Equal(new[] { "ReasoningEffort" }, saves[2]);
            Assert.Equal(new[] { "ReasoningEffort" }, saves[3]);
            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[4]);
            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[5]);
            Assert.Empty(saves[6]);
            Assert.Equal("gpt-6-astra", Config.Model);
            Assert.Equal("High", Config.ReasoningEffort);
        }

        [Theory]
        [InlineData("BtnSaveAiSettings_Click")]
        [InlineData("BtnSavePassword_Click")]
        public void AFailedSave_PutsThisSessionsSettingsBack(string handler)
        {
            // Re-reading config.xml might fail the same way and fall back to
            // defaults (e.g. write tools on), so nothing may change in memory.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";
            Config.AdminPassword = "old";
            Config.WriteToolsEnabled = false;
            Config.EnabledWriteTools = new HashSet<string>();
            string reported = null;
            var save = SettingsForm.SaveSettings;
            var report = SettingsForm.ReportSaveFailed;
            SettingsForm.SaveSettings = (changed, toolsBefore) => "The disk is full.";
            SettingsForm.ReportSaveFailed = (owner, reason) => reported = reason;
            try
            {
                Sta.Run(() =>
                {
                    using (var form = new SettingsForm((CodexAuthService)null))
                    {
                        var flags = BindingFlags.NonPublic | BindingFlags.Instance;
                        typeof(SettingsForm).GetField("_authenticated", flags).SetValue(form, true);
                        Pick(form, "_cmbReasoningEffort", "Low");
                        Field<CheckedListBox>(form, "_clbWriteTools").SetItemChecked(0, true);
                        Field<TextBox>(form, "_txtNewPassword").Text = "new";
                        typeof(SettingsForm).GetMethod(handler, flags).Invoke(form, new object[] { null, EventArgs.Empty });
                    }
                });
            }
            finally
            {
                SettingsForm.SaveSettings = save;
                SettingsForm.ReportSaveFailed = report;
            }

            Assert.Equal("The disk is full.", reported);
            Assert.Equal("High", Config.ReasoningEffort);
            Assert.Equal("old", Config.AdminPassword);
            Assert.False(Config.WriteToolsEnabled);
            Assert.Empty(Config.EnabledWriteTools);
        }

        [Fact]
        public void ARetiredSavedModel_IsSavedOnlyWhenTheAdminPicksOne()
        {
            // Settings shows the replacement on its own; saving it on an unrelated
            // Save could overwrite a model another admin chose since.
            Config.ResetDefaults();
            Config.Model = "gpt-5.5";
            Config.Clock = () => new DateTimeOffset(2026, 10, 15, 0, 0, 0, TimeSpan.Zero);

            var saves = Saves(form =>
            {
                var info = Field<Label>(form, "_lblModelInfo");
                Assert.Equal("gpt-5.6-sol", Shown(form, "_cmbModel"));
                Assert.Contains("Pick it and click Save AI Settings to keep it.", info.Text);

                Tools(form).SetItemChecked(0, false);
                Save(form);
                Assert.Contains("Saved model gpt-5.5 retired", info.Text);
                Pick(form, "_cmbModel", "gpt-5.6-sol");
                Save(form);
                Assert.DoesNotContain("Saved model", info.Text);
            });

            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
            Assert.Equal(new[] { "Model" }, saves[1]);
            Assert.Equal("gpt-5.6-sol", Config.Model);
        }

        [Fact]
        public void SettingsLoadedUnderneathTheDialog_AreNotWrittenBack_AndShowAfterASave()
        {
            // Another dialog's Update Models finished after this one opened and
            // loaded another admin's settings (write tools off, a new model)
            // without refreshing this dialog. Saving an effort change must not
            // write the stale values back, then or on a later Save.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "High";

            var saves = Saves(form =>
            {
                Config.WriteToolsEnabled = false;
                Config.EnabledWriteTools = new HashSet<string>();
                Config.Model = "gpt-6-sol";
                Pick(form, "_cmbReasoningEffort", "Low");
                Save(form);

                Assert.Empty(Tools(form).CheckedItems);
                Assert.Equal("gpt-6-sol", Shown(form, "_cmbModel"));
                Assert.Equal("Low", Shown(form, "_cmbReasoningEffort"));
                Save(form);
            });

            Assert.Equal(new[] { "ReasoningEffort" }, saves[0]);
            Assert.Empty(saves[1]);
            Assert.False(Config.WriteToolsEnabled);
            Assert.Equal("gpt-6-sol", Config.Model);
        }

        [Fact]
        public void SaveAiSettings_SendsTheToolsTheAdminStartedFrom_AndShowsWhatWasSaved()
        {
            // Config merges the tool changes one by one against these (see
            // ConfigTests), even when this dialog's checklist is stale: another
            // dialog's Update Models has loaded tool 0 switched off underneath.
            var t = Config.AllWriteTools;
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            ISet<string> sent = null;
            var save = SettingsForm.SaveSettings;
            SettingsForm.SaveSettings = (changed, toolsBefore) =>
            {
                sent = new HashSet<string>(toolsBefore);
                Config.EnabledWriteTools = new HashSet<string> { t[2], t[3] };   // what the merge saved
                return null;
            };
            try
            {
                Sta.Run(() =>
                {
                    using (var form = new SettingsForm((CodexAuthService)null))
                    {
                        typeof(SettingsForm).GetField("_authenticated", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, true);
                        Config.EnabledWriteTools = new HashSet<string>(t.Skip(1));
                        Tools(form).SetItemChecked(1, false);
                        Save(form);

                        Assert.Equal(new[] { t[2], t[3] }, Tools(form).CheckedItems.Cast<string>());
                    }
                });
            }
            finally
            {
                SettingsForm.SaveSettings = save;
            }

            Assert.Equal(new HashSet<string>(t), sent);
        }

        [Fact]
        public void AfterAReload_EachToolTheAdminDidNotChange_FollowsIt()
        {
            // The admin turns one tool off while another admin turns a different
            // one off: saving keeps both changes.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";

            var saves = Saves(form =>
            {
                Tools(form).SetItemChecked(0, false);
                var others = new HashSet<string>(Config.EnabledWriteTools);
                others.Remove(Config.AllWriteTools[1]);
                Config.EnabledWriteTools = others;   // reloaded: another admin's change
                Reload(form);

                Assert.False(Tools(form).GetItemChecked(0));
                Assert.False(Tools(form).GetItemChecked(1));
                Assert.True(Tools(form).GetItemChecked(2));
                Save(form);
            });

            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
            Assert.DoesNotContain(Config.AllWriteTools[0], Config.EnabledWriteTools);
            Assert.DoesNotContain(Config.AllWriteTools[1], Config.EnabledWriteTools);
            Assert.Equal(Config.AllWriteTools.Length - 2, Config.EnabledWriteTools.Count);
        }

        [Fact]
        public void AnEffortTheModelDoesNotOffer_SurvivesAnUnrelatedSave()
        {
            // gpt-5.5 has no Max, so the dropdown shows Auto; that isn't the admin's change.
            Config.ResetDefaults();
            Config.Model = "gpt-5.5";
            Config.ReasoningEffort = "Max";

            var saves = Saves(form =>
            {
                Assert.Equal("Auto", Shown(form, "_cmbReasoningEffort"));
                Tools(form).SetItemChecked(0, false);
                Save(form);
            });

            Assert.Equal(new[] { "EnabledWriteTools", "WriteToolsEnabled" }, saves[0]);
            Assert.Equal("Max", Config.ReasoningEffort);
        }

        [Theory]
        [InlineData("Max", new[] { "High", "Max" }, new[] { "Model" })]
        [InlineData("Auto", new[] { "Max" }, new[] { "Model", "ReasoningEffort" })]
        public void OnAModelWithoutTheChosenEffort_SaveStoresTheChoice_NotTheAutoShown(string saved, string[] picks, string[] expected)
        {
            // gpt-5.5 has no Max, so the dropdown shows Auto. The admin's choice is
            // what counts and what's saved; gpt-5.5 requests leave the effort out
            // either way, and a later model with Max uses it.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = saved;

            var saves = Saves(form =>
            {
                foreach (var pick in picks) Pick(form, "_cmbReasoningEffort", pick);
                Pick(form, "_cmbModel", "gpt-5.5");
                Assert.Equal("Auto", Shown(form, "_cmbReasoningEffort"));
                Save(form);
            });

            Assert.Equal(expected, saves[0]);
            Assert.Equal("gpt-5.5", Config.Model);
            Assert.Equal("Max", Config.ReasoningEffort);
        }

        [Fact]
        public void AnEffortChoice_SurvivesAModelWithoutIt_AndAReload()
        {
            // Passing through gpt-5.5 shows Auto, which isn't the admin's choice:
            // Update Models must not drop the pending Max.
            Config.ResetDefaults();
            Config.Model = "gpt-6-astra";
            Config.ReasoningEffort = "Auto";

            var saves = Saves(form =>
            {
                Pick(form, "_cmbReasoningEffort", "Max");
                Pick(form, "_cmbModel", "gpt-5.5");
                Assert.Equal("Auto", Shown(form, "_cmbReasoningEffort"));
                Reload(form);
                Pick(form, "_cmbModel", "gpt-6-astra");
                Assert.Equal("Max", Shown(form, "_cmbReasoningEffort"));
                Save(form);
            });

            Assert.Equal(new[] { "ReasoningEffort" }, saves[0]);
            Assert.Equal("Max", Config.ReasoningEffort);
        }

        [Fact]
        public void CatalogLabels_ShowAmpersandsLiterally()
        {
            Sta.Run(() =>
            {
                using (var form = new SettingsForm((CodexAuthService)null))
                {
                    Assert.False(Field<Label>(form, "_lblModelInfo").UseMnemonic);
                    Assert.False(Field<Label>(form, "_lblModelCatalogStatus").UseMnemonic);
                }
            });
        }
    }
}
