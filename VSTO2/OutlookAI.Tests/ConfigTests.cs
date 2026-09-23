using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OutlookAI;
using OutlookAI.Services.Models;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests
{
    [Collection("Config")]
    public class ConfigTests : IDisposable
    {
        // Built-in catalog + a clock before the gpt-5.5 retirement, restored after each test.
        private readonly ConfigStateScope _scope = new ConfigStateScope();

        public void Dispose()
        {
            _scope.Dispose();
        }

        private static (string global, string user) MakeTempPaths()
        {
            var dir = Path.Combine(Path.GetTempPath(),
                "outlookai-config-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            return (Path.Combine(dir, "global.xml"), Path.Combine(dir, "user.xml"));
        }

        [Fact]
        public void LoadConfigFromPaths_UsesV2Defaults_WhenFilesAreMissing()
        {
            var (g, u) = MakeTempPaths();
            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("admin", Config.AdminPassword);
            Assert.Equal(@"C:\ProgramData\OutlookAI\auth.json", Config.CodexAuthPath);
            // Default model = the catalog's top listed, non-retired model.
            Assert.Equal("gpt-6-astra", Config.DefaultModel);
            Assert.Equal("gpt-6-astra", Config.Model);
            Assert.Equal("gpt-realtime-1.5", Config.VoiceModel);
            Assert.Equal("", Config.ModelCatalogClientVersion);
        }

        [Fact]
        public void LoadConfigFromPaths_PerUserOverridesAdminPasswordOnly()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<AdminPassword>server</AdminPassword>"
                + "<CodexAuthPath>C:\\ProgramData\\OutlookAI\\auth.json</CodexAuthPath>"
                + "<Model>gpt-5.5</Model>"
                + "<VoiceModel>gpt-realtime-1.5</VoiceModel>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<AdminPassword>userpass</AdminPassword>"
                + "<Model>claude-opus-4-6</Model>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("userpass", Config.AdminPassword);
            // Server-authoritative Model is not overridden by per-user;
            // unknown Claude-era model names also do not override it.
            Assert.Equal("gpt-5.5", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresLegacyV1Fields()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<ApiKey>anthropic-key</ApiKey>"
                + "<OpenAIApiKey>openai-key</OpenAIApiKey>"
                + "<WhisperModel>whisper-1</WhisperModel>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal(Config.DefaultModel, Config.Model);
            Assert.Equal("gpt-realtime-1.5", Config.VoiceModel);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesReasoningEffortFromGlobal()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("High", Config.ReasoningEffort);
        }

        [Fact]
        public void LoadConfigFromPaths_UserOverridesReasoningEffortAndWriteTools()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<ReasoningEffort>Medium</ReasoningEffort>"
                + "<WriteToolsEnabled>true</WriteToolsEnabled>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<ReasoningEffort>Low</ReasoningEffort>"
                + "<WriteToolsEnabled>false</WriteToolsEnabled>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("Low", Config.ReasoningEffort);
            Assert.False(Config.WriteToolsEnabled);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresUnknownReasoningEffort()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>Extreme</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            // Unknown value -> fall back to the default, Auto.
            Assert.Equal("Auto", Config.ReasoningEffort);
        }

        [Fact]
        public void ReasoningEffortsForModel_ModelOutsideTheCatalog_OffersOnlyAuto()
        {
            Assert.Equal(new[] { "Auto" }, Config.ReasoningEffortsForModel("gpt-4.1-nano"));
            Assert.Equal(new[] { "Auto" }, Config.ReasoningEffortsForModel("gpt-5.5-pro"));
            Assert.Contains("High", Config.ReasoningEffortsForModel("gpt-5.5"));
        }

        [Fact]
        public void ReasoningEffortsForModel_Gpt55_ExcludesMinimal_IncludesXHigh()
        {
            // Backend ground truth: 'minimal' and 'max' are rejected for gpt-5.5.
            var efforts = Config.ReasoningEffortsForModel("gpt-5.5");
            Assert.DoesNotContain("Minimal", efforts);
            Assert.Contains("XHigh", efforts);
            Assert.Equal(new[] { "Auto", "Low", "Medium", "High", "XHigh" }, efforts);
        }

        [Fact]
        public void ReasoningEffortsForModel_Gpt6Astra_OffersMax_NotTheClientOnlyUltra()
        {
            var efforts = Config.ReasoningEffortsForModel("gpt-6-astra");
            Assert.Contains("Max", efforts);
            Assert.DoesNotContain("Ultra", efforts);
        }

        [Fact]
        public void ReasoningEffortsForModel_FollowsARefreshedCatalog()
        {
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("gpt-7", 1, new[] { "low", "extreme" }));

            Assert.Equal(new[] { "Auto", "Low" }, Config.ReasoningEffortsForModel("gpt-7"));
            Config.ModelCatalog = new ModelCatalog(Config.ModelCatalog.Models, new[] { "low", "extreme" },
                ModelCatalog.SourceChatGpt, TestCatalogs.FixedNow, "0.160.0");
            Assert.Equal(new[] { "Auto", "Low", "Extreme" }, Config.ReasoningEffortsForModel("gpt-7"));
        }

        [Fact]
        public void LoadConfigFromPaths_RecordsEveryModelNamedInConfig_EvenRejectedOnes()
        {
            // Update Models keeps these alive when a filtered list omits them.
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(g, "<Config><Model> gpt-7 </Model></Config>");
            File.WriteAllText(s, "<Config><Model>GPT-6-SOL</Model></Config>");
            File.WriteAllText(u, "<Config><Model>gpt-5.5</Model></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, s, u);

                // Load order: global, per-user, then the server-wide file, which wins.
                Assert.Equal(new[] { "gpt-7", "gpt-5.5", "GPT-6-SOL" }, Config.ModelsNamedInConfig);
                Assert.Equal("gpt-6-sol", Config.Model);
            }
            finally
            {
                File.Delete(g);
                File.Delete(s);
                File.Delete(u);
            }
        }

        [Fact]
        public void NotifyAiSettingsChanged_RunsEveryHandler_EvenIfOneThrows()
        {
            var calls = 0;
            EventHandler failing = (s, e) => { calls++; throw new InvalidOperationException("pane disposed"); };
            EventHandler working = (s, e) => calls++;
            Config.AiSettingsChanged += failing;
            Config.AiSettingsChanged += working;
            try
            {
                Config.NotifyAiSettingsChanged();
                Assert.Equal(2, calls);
            }
            finally
            {
                Config.AiSettingsChanged -= failing;
                Config.AiSettingsChanged -= working;
            }
        }

        [Fact]
        public void Defaults_AreV2()
        {
            var (g, u) = MakeTempPaths();
            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("Auto", Config.ReasoningEffort);
            Assert.True(Config.WriteToolsEnabled);
            // Default: all four write tools enabled.
            Assert.Equal(4, Config.EnabledWriteTools.Count);
            Assert.Contains("outlook_create_draft", Config.EnabledWriteTools);
            Assert.Contains("outlook_mark_as_read", Config.EnabledWriteTools);
            Assert.Contains("outlook_flag_message", Config.EnabledWriteTools);
            Assert.Contains("outlook_set_category", Config.EnabledWriteTools);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesEnabledWriteToolsFromCSV()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<EnabledWriteTools>outlook_create_draft, outlook_set_category</EnabledWriteTools>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal(2, Config.EnabledWriteTools.Count);
            Assert.Contains("outlook_create_draft", Config.EnabledWriteTools);
            Assert.Contains("outlook_set_category", Config.EnabledWriteTools);
            Assert.DoesNotContain("outlook_mark_as_read", Config.EnabledWriteTools);
        }

        [Fact]
        public void LoadConfigFromPaths_FiltersUnknownWriteToolNames()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<EnabledWriteTools>outlook_create_draft,outlook_send_now,bogus_tool</EnabledWriteTools>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            // Unknown names are silently dropped (forward-compat with admin
            // configs that reference future or removed tools).
            Assert.Single(Config.EnabledWriteTools);
            Assert.Contains("outlook_create_draft", Config.EnabledWriteTools);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesModelFromUserOverride_WhenInCatalog()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>gpt-5.5</Model></Config>");
            File.WriteAllText(u, "<Config><Model>gpt-6-sol</Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("gpt-6-sol", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_IgnoresModelNotInCatalog()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>nonexistent-gpt-9</Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            // Falls back to the catalog default since the requested model
            // isn't in the catalog.
            Assert.Equal(Config.DefaultModel, Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_AcceptsModelsFromARefreshedCatalog()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>gpt-7</Model></Config>");

            Config.LoadConfigFromPaths(g, u);
            Assert.NotEqual("gpt-7", Config.Model);   // built-in catalog doesn't know it

            // gpt-7 is deliberately not the catalog default, so only the config
            // value can put it there.
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("gpt-top", 1, new[] { "low" }),
                TestCatalogs.Entry("gpt-7", 2, new[] { "low" }),
                TestCatalogs.Entry("gpt-reserve", 3, new[] { "low" }, listed: false));
            Config.LoadConfigFromPaths(g, u);
            Assert.Equal("gpt-7", Config.Model);

            // Hidden catalog models are still valid when set explicitly.
            File.WriteAllText(g, "<Config><Model>gpt-reserve</Model></Config>");
            Config.LoadConfigFromPaths(g, u);
            Assert.Equal("gpt-reserve", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_MatchesModelCaseInsensitively_AndStoresTheCanonicalSlug()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><Model>  GPT-5.6-SOL </Model></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("gpt-5.6-sol", Config.Model);
        }

        [Theory]
        [InlineData("max", "Max")]
        [InlineData("XHIGH", "XHigh")]
        [InlineData("medium", "Medium")]
        public void LoadConfigFromPaths_AcceptsCatalogEfforts_WithCanonicalCasing(string raw, string expected)
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>" + raw + "</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal(expected, Config.ReasoningEffort);
        }

        [Theory]
        [InlineData("Ultra")]     // Codex client-only mode; the server rejects it
        [InlineData("Minimal")]   // no current model accepts it
        public void LoadConfigFromPaths_IgnoresEffortsNoModelOffers(string raw)
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>" + raw + "</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("Auto", Config.ReasoningEffort);
        }

        [Theory]
        [InlineData("None")]   // what Auto was called before v2.2.1; configs still say it
        [InlineData("none")]
        [InlineData("auto")]
        public void LoadConfigFromPaths_ReadsTheLegacyNoneAsAuto(string raw)
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");
            File.WriteAllText(u, "<Config><ReasoningEffort>" + raw + "</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("Auto", Config.ReasoningEffort);
        }

        [Theory]
        [InlineData("Auto", "None")]   // the name v2.2.0 and older read as "omit"
        [InlineData("High", "High")]
        public void BuildSavedConfig_StoresAutoAsNone_SoOlderVersionsReadItTheSame(string effort, string stored)
        {
            Config.ReasoningEffort = effort;

            var saved = Config.BuildSavedConfig();

            Assert.Equal(stored, saved.Root.Element("ReasoningEffort").Value);
        }

        [Fact]
        public void BuildSavedConfig_AutoRoundTripsOverAHigherSharedDefault()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");
            Config.ReasoningEffort = "Auto";
            Config.BuildSavedConfig().Save(u);

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("Auto", Config.ReasoningEffort);
        }

        [Fact]
        public void LoadConfigFromPaths_ModelCatalogClientVersion_IsServerAuthoritative()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><ModelCatalogClientVersion> 0.160.0 </ModelCatalogClientVersion></Config>");
            File.WriteAllText(u, "<Config><ModelCatalogClientVersion>0.170.0</ModelCatalogClientVersion></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("0.160.0", Config.ModelCatalogClientVersion);
        }

        [Fact]
        public void EffectiveModel_SwitchesToTheUpgradeTarget_OnceTheModelRetires()
        {
            Config.Model = "gpt-5.5";
            var retirement = new DateTimeOffset(2026, 10, 14, 19, 0, 0, TimeSpan.Zero);

            Config.Clock = () => retirement.AddMinutes(-1);
            Assert.Equal("gpt-5.5", Config.EffectiveModel);

            Config.Clock = () => retirement;
            Assert.Equal("gpt-5.6-sol", Config.EffectiveModel);
            Assert.Equal("gpt-5.5", Config.Model);   // the saved choice is untouched
        }

        [Fact]
        public void EffectiveModel_UsesTheDefault_WhenTheModelLeavesTheCatalog()
        {
            Config.Model = "gpt-5.5";
            Config.ModelCatalog = TestCatalogs.Catalog(
                TestCatalogs.Entry("gpt-7", 1, new[] { "low" }),
                TestCatalogs.Entry("gpt-6-sol", 2, new[] { "low" }));

            Assert.Equal("gpt-7", Config.EffectiveModel);
            Assert.Equal("gpt-7", Config.DefaultModel);
        }

        [Fact]
        public void ResetDefaults_KeepsTheLoadedCatalog()
        {
            var refreshed = TestCatalogs.Catalog(TestCatalogs.Entry("gpt-7", 1, new[] { "low" }));
            Config.ModelCatalog = refreshed;

            Config.ResetDefaults();

            Assert.Same(refreshed, Config.ModelCatalog);
            Assert.Equal("gpt-7", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_SharedDefaultsAppliedWhenUserAbsent()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(s, "<Config><ReasoningEffort>Medium</ReasoningEffort></Config>");
            // No user file
            try
            {
                Config.LoadConfigFromPaths(g, s, u);
                Assert.Equal("Medium", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(s)) File.Delete(s);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_ServerWideSettingsBeatAPerUserFile()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(s, "<Config><ReasoningEffort>Medium</ReasoningEffort></Config>");
            File.WriteAllText(u, "<Config><ReasoningEffort>Low</ReasoningEffort></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, s, u);
                Assert.Equal("Medium", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(s)) File.Delete(s);
                if (File.Exists(u)) File.Delete(u);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_SharedConfigPathNull_TreatedAsAbsent()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(u, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal("High", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(u)) File.Delete(u);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_GlobalAndSharedAndUser_SharedBeatsUserBeatsGlobal()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(g, "<Config><ReasoningEffort>High</ReasoningEffort></Config>");
            File.WriteAllText(s, "<Config><ReasoningEffort>Medium</ReasoningEffort></Config>");
            File.WriteAllText(u, "<Config><ReasoningEffort>Low</ReasoningEffort></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, s, u);
                Assert.Equal("Medium", Config.ReasoningEffort);
            }
            finally
            {
                if (File.Exists(g)) File.Delete(g);
                if (File.Exists(s)) File.Delete(s);
                if (File.Exists(u)) File.Delete(u);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_APerUserCopyCantShadowServerWideSettings()
        {
            // What Settings wrote per user before v2.2.2, then the admin changed everything server-wide.
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "shared.xml");
            File.WriteAllText(g, "<Config><AdminPassword>installer</AdminPassword></Config>");
            File.WriteAllText(s, "<Config><AdminPassword>rotated</AdminPassword><Model>gpt-6-sol</Model>"
                + "<ReasoningEffort>High</ReasoningEffort><WriteToolsEnabled>true</WriteToolsEnabled>"
                + "<EnabledWriteTools>outlook_create_draft</EnabledWriteTools></Config>");
            File.WriteAllText(u, "<Config><AdminPassword>old</AdminPassword><Model>gpt-5.5</Model>"
                + "<ReasoningEffort>Low</ReasoningEffort><WriteToolsEnabled>false</WriteToolsEnabled>"
                + "<EnabledWriteTools>outlook_mark_as_read</EnabledWriteTools></Config>");

            Config.LoadConfigFromPaths(g, s, u);

            Assert.Equal("rotated", Config.AdminPassword);
            Assert.Equal("gpt-6-sol", Config.Model);
            Assert.Equal("High", Config.ReasoningEffort);
            Assert.True(Config.WriteToolsEnabled);
            Assert.Equal(new[] { "outlook_create_draft" }, Config.EnabledWriteTools);
        }

        [Theory]
        [InlineData("<EnabledWriteTools></EnabledWriteTools>")]
        [InlineData("<EnabledWriteTools />")]
        public void LoadConfigFromPaths_AnEmptyToolList_MeansEveryToolUnchecked(string element)
        {
            // What Settings saves when every write tool is unchecked; it must not
            // bring back an earlier list for the next Save to switch on again.
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "shared.xml");
            File.WriteAllText(g, "<Config><WriteToolsEnabled>true</WriteToolsEnabled><EnabledWriteTools>outlook_create_draft</EnabledWriteTools></Config>");
            File.WriteAllText(s, "<Config><WriteToolsEnabled>false</WriteToolsEnabled>" + element + "</Config>");

            Config.LoadConfigFromPaths(g, s, u);

            Assert.False(Config.WriteToolsEnabled);
            Assert.Empty(Config.EnabledWriteTools);
        }

        [Fact]
        public void LoadConfigFromPaths_APartialServerWideFile_KeepsThePerUserValuesItDoesNotSet()
        {
            // A hand-made server-wide file with just a model must not reset the
            // rest (e.g. re-enable write tools or restore the installer's password).
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "shared.xml");
            File.WriteAllText(g, "<Config><AdminPassword>installer</AdminPassword></Config>");
            File.WriteAllText(s, "<Config><Model>gpt-6-sol</Model></Config>");
            File.WriteAllText(u, "<Config><AdminPassword>rotated</AdminPassword><Model>gpt-5.5</Model>"
                + "<WriteToolsEnabled>false</WriteToolsEnabled><EnabledWriteTools></EnabledWriteTools></Config>");

            Config.LoadConfigFromPaths(g, s, u);

            Assert.Equal("gpt-6-sol", Config.Model);
            Assert.Equal("rotated", Config.AdminPassword);
            Assert.False(Config.WriteToolsEnabled);
        }

        [Theory]
        [InlineData(null)]                     // never saved server-wide (e.g. no Settings save since v2.1.1)
        [InlineData("<Config><Model>")]       // unreadable
        public void LoadConfigFromPaths_WithoutReadableServerWideSettings_APerUserFileStillApplies(string shared)
        {
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "shared.xml");
            if (shared != null) File.WriteAllText(s, shared);
            File.WriteAllText(g, "<Config><AdminPassword>installer</AdminPassword></Config>");
            File.WriteAllText(u, "<Config><AdminPassword>rotated</AdminPassword><ReasoningEffort>Low</ReasoningEffort></Config>");

            Config.LoadConfigFromPaths(g, s, u);

            Assert.Equal("rotated", Config.AdminPassword);
            Assert.Equal("Low", Config.ReasoningEffort);
        }

        [Fact]
        public void SaveSettingsTo_SavesForEveryUser_EvenOneWithAnOldPerUserCopy()
        {
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "ProgramData", "config.xml");
            File.WriteAllText(u, "<Config><AdminPassword>old</AdminPassword><Model>gpt-5.5</Model><ReasoningEffort>Low</ReasoningEffort>"
                + "<WriteToolsEnabled>false</WriteToolsEnabled><EnabledWriteTools></EnabledWriteTools></Config>");
            Config.ResetDefaults();   // all four write tools on
            Config.Model = "gpt-6-sol";
            Config.ReasoningEffort = "High";
            Config.AdminPassword = "new";

            Assert.Null(Config.SaveSettingsTo(s, new[] { "Model", "ReasoningEffort" }));
            Assert.Equal(Config.SavedSettingNames, ElementNames(s));   // the first save writes all of them

            Config.ResetDefaults();
            Config.LoadConfigFromPaths(g, s, u);   // that user's next Outlook start

            Assert.Equal("gpt-6-sol", Config.Model);
            Assert.Equal("High", Config.ReasoningEffort);
            Assert.Equal("new", Config.AdminPassword);
            Assert.True(Config.WriteToolsEnabled);
        }

        [Fact]
        public void SaveSettingsTo_WritesOnlyTheChangedSettings_OverWhatAnotherSessionSaved()
        {
            // This Outlook loaded the settings before another admin changed the password and effort.
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            File.WriteAllText(s, "<Config><AdminPassword>rotated</AdminPassword><Model>gpt-5.5</Model><ReasoningEffort>High</ReasoningEffort>"
                + "<WriteToolsEnabled>true</WriteToolsEnabled><EnabledWriteTools>outlook_create_draft</EnabledWriteTools></Config>");
            Config.ResetDefaults();
            Config.AdminPassword = "stale";
            Config.ReasoningEffort = "Low";
            Config.Model = "gpt-6-sol";

            Assert.Null(Config.SaveSettingsTo(s, new[] { "Model" }));

            var root = XDocument.Load(s).Root;
            Assert.Equal("gpt-6-sol", root.Element("Model").Value);
            Assert.Equal("rotated", root.Element("AdminPassword").Value);
            Assert.Equal("High", root.Element("ReasoningEffort").Value);
            Assert.Equal("outlook_create_draft", root.Element("EnabledWriteTools").Value);
        }

        [Fact]
        public void SaveSettingsTo_MergesWriteToolsOneByOne_WithWhatAnotherSessionSaved()
        {
            // This session started from tools 0-2 on, then switched 1 off and 3
            // on; another admin has since switched 0 off. All three changes stay.
            var t = Config.AllWriteTools;
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            File.WriteAllText(s, "<Config><WriteToolsEnabled>true</WriteToolsEnabled><EnabledWriteTools>"
                + t[1] + "," + t[2] + "</EnabledWriteTools></Config>");
            Config.ResetDefaults();
            Config.EnabledWriteTools = new HashSet<string> { t[0], t[2], t[3] };

            Assert.Null(Config.SaveSettingsTo(s, new[] { "EnabledWriteTools", "WriteToolsEnabled" }, new HashSet<string> { t[0], t[1], t[2] }));

            var root = XDocument.Load(s).Root;
            Assert.Equal(new[] { t[2], t[3] }, root.Element("EnabledWriteTools").Value.Split(',').OrderBy(x => Array.IndexOf(t, x)));
            Assert.Equal("true", root.Element("WriteToolsEnabled").Value);
            Assert.Equal(new[] { t[2], t[3] }, Config.EnabledWriteTools.OrderBy(x => Array.IndexOf(t, x)));   // this session uses what was saved
            Assert.True(Config.WriteToolsEnabled);
        }

        [Fact]
        public void SaveSettingsTo_ToolsSwitchedOffHere_KeepWritesOff_WhenAnotherSessionSwitchedThemOff()
        {
            var t = Config.AllWriteTools;
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            File.WriteAllText(s, "<Config><WriteToolsEnabled>false</WriteToolsEnabled><EnabledWriteTools>"
                + string.Join(",", t) + "</EnabledWriteTools></Config>");
            Config.ResetDefaults();
            Config.EnabledWriteTools = new HashSet<string>(t.Skip(1));

            Assert.Null(Config.SaveSettingsTo(s, new[] { "EnabledWriteTools", "WriteToolsEnabled" }, new HashSet<string>(t)));

            var root = XDocument.Load(s).Root;
            Assert.Equal("false", root.Element("WriteToolsEnabled").Value);
            Assert.Equal("", root.Element("EnabledWriteTools").Value);
            Assert.False(Config.WriteToolsEnabled);
        }

        [Fact]
        public void SaveSettingsTo_ASavedToolListAnotherLayerSwitchesOff_IsNotSwitchedOnByOneTool()
        {
            // The server-wide file lists every tool but has no switch, and the
            // Program Files config switches writes off, so none are in effect.
            // Checking one tool in Settings must enable just that one.
            var t = Config.AllWriteTools;
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "shared.xml");
            File.WriteAllText(g, "<Config><WriteToolsEnabled>false</WriteToolsEnabled></Config>");
            File.WriteAllText(s, "<Config><EnabledWriteTools>" + string.Join(",", t) + "</EnabledWriteTools></Config>");
            Config.LoadConfigFromPaths(g, s, u);
            Assert.False(Config.WriteToolsEnabled);

            Config.EnabledWriteTools = new HashSet<string> { t[0] };   // what Save sets for the one checked tool
            Config.WriteToolsEnabled = true;
            Assert.Null(Config.SaveSettingsTo(s, new[] { "EnabledWriteTools", "WriteToolsEnabled" }, new HashSet<string>()));

            Config.ResetDefaults();
            Config.LoadConfigFromPaths(g, s, u);
            Assert.True(Config.WriteToolsEnabled);
            Assert.Equal(new[] { t[0] }, Config.EnabledWriteTools);
        }

        [Fact]
        public void SaveSettingsTo_WithoutASavedToolList_SavesThisSessionsTools()
        {
            var t = Config.AllWriteTools;
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            Config.ResetDefaults();
            Config.EnabledWriteTools = new HashSet<string>(t.Skip(1));

            Assert.Null(Config.SaveSettingsTo(s, new[] { "EnabledWriteTools", "WriteToolsEnabled" }, new HashSet<string>(t)));

            Assert.Equal(t.Skip(1), XDocument.Load(s).Root.Element("EnabledWriteTools").Value.Split(',').OrderBy(x => Array.IndexOf(t, x)));
        }

        [Fact]
        public void SaveSettingsTo_FillsInSettingsMissingFromTheFile_AndLeavesNothingElseBehind()
        {
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            File.WriteAllText(s, "<Config><Model>gpt-5.5</Model><AdminPassword> </AdminPassword><VoiceModel>hand-added</VoiceModel></Config>");
            Config.ResetDefaults();
            Config.AdminPassword = "mine";
            Config.ReasoningEffort = "High";

            Assert.Null(Config.SaveSettingsTo(s, new[] { "ReasoningEffort" }));

            var root = XDocument.Load(s).Root;
            Assert.Equal("gpt-5.5", root.Element("Model").Value);
            Assert.Equal("hand-added", root.Element("VoiceModel").Value);
            Assert.Equal("High", root.Element("ReasoningEffort").Value);
            Assert.Equal("mine", root.Element("AdminPassword").Value);
            Assert.All(Config.SavedSettingNames, name => Assert.NotNull(root.Element(name)));
            Assert.Equal(new[] { s }, Directory.GetFileSystemEntries(Path.GetDirectoryName(s), "config.xml*"));   // no temp or lock files
        }

        [Fact]
        public void SaveSettingsTo_WaitsForAnotherSessionsSave_ThenSaysSo()
        {
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            using (new FileStream(s + ".lock", FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                var error = Config.SaveSettingsTo(s, new[] { "Model" }, lockTimeout: TimeSpan.FromMilliseconds(200));

                Assert.Contains("Try again", error);
                Assert.False(File.Exists(s));
            }
        }

        [Fact]
        public void LoadConfigFromPaths_WaitsBrieflyForAServerWideFileAnotherSessionIsSaving()
        {
            // Otherwise this whole session would run on the old per-user copy.
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "shared.xml");
            File.WriteAllText(s, "<Config><Model>gpt-6-sol</Model></Config>");
            File.WriteAllText(u, "<Config><Model>gpt-5.5</Model></Config>");
            var busy = new FileStream(s, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var release = System.Threading.Tasks.Task.Delay(300).ContinueWith(_ => busy.Dispose());

            Config.LoadConfigFromPaths(g, s, u);

            release.Wait();
            Assert.Equal("gpt-6-sol", Config.Model);
        }

        [Fact]
        public void SaveSettingsTo_WhenTheFileCantBeRead_FailsEvenIfItFreesUpBeforeTheSwap()
        {
            // Another process holds config.xml past the read's retries (~1 s), then
            // lets go while the swap still retries. Writing every setting from this
            // session then would undo other admins' changes.
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            const string saved = "<Config><AdminPassword>rotated</AdminPassword><WriteToolsEnabled>false</WriteToolsEnabled></Config>";
            File.WriteAllText(s, saved);
            Config.ResetDefaults();
            Config.ReasoningEffort = "High";
            var busy = new FileStream(s, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var release = System.Threading.Tasks.Task.Delay(1500).ContinueWith(_ => busy.Dispose());

            var error = Config.SaveSettingsTo(s, new[] { "ReasoningEffort" });

            release.Wait();
            Assert.False(string.IsNullOrEmpty(error));
            Assert.Equal(saved, File.ReadAllText(s));
        }

        [Fact]
        public void SaveSettingsTo_AFileThatIsNotValidXml_IsReportedNotReplaced()
        {
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            File.WriteAllText(s, "<Config><Model>");

            var error = Config.SaveSettingsTo(s, new[] { "Model" });

            Assert.Contains("not valid XML", error);
            Assert.Equal("<Config><Model>", File.ReadAllText(s));
        }

        [Fact]
        public void SaveSettingsTo_WaitsForAnotherSessionReadingTheFile()
        {
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            File.WriteAllText(s, "<Config><ReasoningEffort>Low</ReasoningEffort></Config>");
            Config.ResetDefaults();
            Config.ReasoningEffort = "High";
            var reader = new FileStream(s, FileMode.Open, FileAccess.Read, FileShare.Read);
            var release = System.Threading.Tasks.Task.Delay(300).ContinueWith(_ => reader.Dispose());

            var error = Config.SaveSettingsTo(s, new[] { "ReasoningEffort" });

            release.Wait();
            Assert.Null(error);
            Assert.Equal("High", XDocument.Load(s).Root.Element("ReasoningEffort").Value);
        }

        [Fact]
        public void ReloadConfigFiles_KeepsWhatThisSessionHas_WhileASavedFileCantBeRead()
        {
            // Falling back to the other layers would e.g. switch write tools back on.
            var (g, u) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "shared.xml");
            File.WriteAllText(s, "<Config><WriteToolsEnabled>false</WriteToolsEnabled><EnabledWriteTools></EnabledWriteTools></Config>");
            Config.LoadConfigFromPaths(g, s, u);

            using (new FileStream(s, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                Config.ReloadConfigFilesFrom(g, s, u);   // gives up after ~1 s
            }
            Assert.False(Config.WriteToolsEnabled);

            File.WriteAllText(s, "<Config><WriteToolsEnabled>true</WriteToolsEnabled></Config>");
            Config.ReloadConfigFilesFrom(g, s, u);
            Assert.True(Config.WriteToolsEnabled);
        }

        [Fact]
        public void ReloadConfigFiles_KeepsWhatThisSessionHas_WhenTheSavedFilesFolderIsDenied()
        {
            // File.Exists then says false, as if there were no file; the other
            // layers would e.g. switch write tools back on.
            var (g, u) = MakeTempPaths();
            var dir = Path.Combine(Path.GetDirectoryName(g), "ProgramData");
            Directory.CreateDirectory(dir);
            var s = Path.Combine(dir, "config.xml");
            File.WriteAllText(s, "<Config><WriteToolsEnabled>false</WriteToolsEnabled><EnabledWriteTools></EnabledWriteTools></Config>");
            Config.LoadConfigFromPaths(g, s, u);

            var deny = new System.Security.AccessControl.FileSystemAccessRule(
                System.Security.Principal.WindowsIdentity.GetCurrent().User,
                System.Security.AccessControl.FileSystemRights.ListDirectory | System.Security.AccessControl.FileSystemRights.ReadAttributes,
                System.Security.AccessControl.InheritanceFlags.ContainerInherit | System.Security.AccessControl.InheritanceFlags.ObjectInherit,
                System.Security.AccessControl.PropagationFlags.None,
                System.Security.AccessControl.AccessControlType.Deny);
            var acl = Directory.GetAccessControl(dir);
            acl.AddAccessRule(deny);
            Directory.SetAccessControl(dir, acl);
            try
            {
                Assert.False(File.Exists(s));
                Config.ReloadConfigFilesFrom(g, s, u);
            }
            finally
            {
                acl.RemoveAccessRule(deny);
                Directory.SetAccessControl(dir, acl);
            }

            Assert.False(Config.WriteToolsEnabled);
        }

        [Fact]
        public void Saving_NeverWritesThePerUserFile()
        {
            // Settings saves for every user; the per-user path is only passed to the loaders.
            var uses = File.ReadAllLines(RepoFiles.Find("VSTO2", "OutlookAI", "Config.cs"))
                .Where(line => line.Contains("UserConfigFilePath") && !line.Contains("private static readonly string UserConfigFilePath"))
                .ToArray();

            Assert.NotEmpty(uses);
            Assert.All(uses, line => Assert.Contains("(GlobalConfigFilePath, SharedConfigFilePath, UserConfigFilePath);", line));
        }

        [Fact]
        public void SaveSettingsTo_ReportsWhyItCouldNotSave()
        {
            var (g, _) = MakeTempPaths();
            var s = Path.Combine(Path.GetDirectoryName(g), "config.xml");
            Directory.CreateDirectory(s);   // a folder where the file should go

            Assert.False(string.IsNullOrEmpty(Config.SaveSettingsTo(s, new[] { "Model" })));
        }

        private static string[] ElementNames(string path)
        {
            return XDocument.Load(path).Root.Elements().Select(e => e.Name.LocalName).ToArray();
        }

        [Fact]
        public void MaxBulkExportRows_DefaultsTo2000()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
            Assert.Equal(2000, Config.MaxBulkExportRows);
        }

        [Fact]
        public void MaxBulkExportRows_LoadsFromGlobalConfig()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(g, "<Config><MaxBulkExportRows>500</MaxBulkExportRows></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(500, Config.MaxBulkExportRows);
            }
            finally { if (File.Exists(g)) File.Delete(g); }
        }

        [Fact]
        public void MaxBulkExportRows_ClampsToFloorAndCeiling()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(g, "<Config><MaxBulkExportRows>0</MaxBulkExportRows></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(1, Config.MaxBulkExportRows);   // floor

                File.WriteAllText(g, "<Config><MaxBulkExportRows>999999</MaxBulkExportRows></Config>");
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                // Ceiling shares the interactive export cap (#12.1) so neither
                // Excel path can exceed BulkExportRowCap.Max (10,000).
                Assert.Equal(10000, Config.MaxBulkExportRows);  // ceiling
            }
            finally { if (File.Exists(g)) File.Delete(g); }
        }

        [Fact]
        public void MaxBulkExportRows_NotUserOverridable()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(g, "<Config><MaxBulkExportRows>750</MaxBulkExportRows></Config>");
            File.WriteAllText(u, "<Config><MaxBulkExportRows>3000</MaxBulkExportRows></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(750, Config.MaxBulkExportRows);  // user value ignored
            }
            finally { if (File.Exists(g)) File.Delete(g); if (File.Exists(u)) File.Delete(u); }
        }
    }
}
