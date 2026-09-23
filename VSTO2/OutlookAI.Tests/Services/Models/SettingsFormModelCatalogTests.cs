using System;
using System.Globalization;
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
