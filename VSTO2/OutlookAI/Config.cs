using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OutlookAI.Services.Models;

namespace OutlookAI
{
    public static class Config
    {
        // ============================================================
        // CONFIGURATION DEFAULTS (v2 - ChatGPT OAuth)
        // Server-authoritative fields (CodexAuthPath, Model) load from
        // defaults -> global config (Program Files). Per-user AppData
        // config may override only AdminPassword. Legacy v1 elements
        // (ApiKey, OpenAIApiKey, WhisperModel, TranscribeModel, MaxTokens,
        // and Claude model names) are ignored if encountered.
        // ============================================================

        public const string DefaultVoiceModel = "gpt-realtime-1.5";
        public const string DefaultCodexAuthPath = @"C:\ProgramData\OutlookAI\auth.json";
        public const string DefaultReasoningEffort = ReasoningEffortNames.Auto;
        public const bool DefaultWriteToolsEnabled = true;
        public const int DefaultMaxBulkExportRows = 2000;
        private const int MinBulkExportRows = 1;
        // Shared with the interactive export path so the two Excel exporters
        // can never produce workbooks of differing maximum size (#12.1).
        private const int MaxBulkExportRowsCeiling = Services.Tools.BulkExportRowCap.Max;

        /// <summary>
        /// Models and their reasoning efforts. Built-in until an admin runs
        /// Settings → Update Models, then the cached models.json, loaded by
        /// <see cref="LoadConfig"/> before the config layers (which validate
        /// against it). <see cref="ResetDefaults"/> leaves it alone.
        /// </summary>
        public static ModelCatalog ModelCatalog { get; set; } = BuiltInModelCatalog.Instance;

        // Test seam for retirement-date resolution.
        internal static Func<DateTimeOffset> Clock { get; set; } = () => DateTimeOffset.UtcNow;

        public static string AdminPassword { get; set; } = "admin";
        public static string CodexAuthPath { get; set; } = DefaultCodexAuthPath;
        // After ModelCatalog/Clock: static initializers run in textual order.
        public static string Model { get; set; } = DefaultModel;
        public static string VoiceModel { get; set; } = DefaultVoiceModel;

        /// <summary>The catalog's first listed model that hasn't retired.</summary>
        public static string DefaultModel => ModelCatalog.DefaultModelAt(Clock());

        /// <summary>
        /// The model requests are sent with: <see cref="Model"/>, unless the
        /// catalog says it has retired (then its upgrade target) or no longer
        /// offers it (then <see cref="DefaultModel"/>). The saved
        /// <see cref="Model"/> is never rewritten.
        /// </summary>
        public static string EffectiveModel => ModelCatalog.ResolveEffectiveModel(Model, Clock());

        /// <summary>
        /// Optional client_version for the model-catalog request (Program Files
        /// config.xml only). Blank = track the latest Codex CLI release.
        /// </summary>
        public static string ModelCatalogClientVersion { get; set; } = "";

        /// <summary>
        /// Default reasoning effort sent to the Codex backend on each turn:
        /// "Auto" or an effort some catalog model offers. "Auto" (formerly
        /// "None") omits the reasoning block so the model's default applies.
        /// Per-turn overrides via <c>ConversationContext.ReasoningEffortOverride</c>.
        /// </summary>
        public static string ReasoningEffort { get; set; } = DefaultReasoningEffort;

        /// <summary>
        /// Master switch for the four safe-write Outlook tools (create_draft,
        /// mark_as_read, flag_message, set_category). When false, the tool
        /// catalog sent to the model only includes the read tools. When true,
        /// the per-tool set <see cref="EnabledWriteTools"/> determines which
        /// individual writes are surfaced.
        /// </summary>
        public static bool WriteToolsEnabled { get; set; } = DefaultWriteToolsEnabled;

        /// <summary>
        /// Hard ceiling on rows collected by outlook_export_search_results.
        /// Server-authoritative (global config only); not user-overridable.
        /// Bounds runtime/memory on large mailboxes. Clamped to
        /// [MinBulkExportRows, MaxBulkExportRowsCeiling] on load.
        /// </summary>
        public static int MaxBulkExportRows { get; set; } = DefaultMaxBulkExportRows;

        /// <summary>
        /// Full set of write-tool names supported by Phase 2. Used both as
        /// the SettingsForm option list and as the default for
        /// <see cref="EnabledWriteTools"/>.
        /// </summary>
        public static readonly string[] AllWriteTools =
        {
            "outlook_create_draft",
            "outlook_mark_as_read",
            "outlook_flag_message",
            "outlook_set_category"
        };

        /// <summary>
        /// Currently-enabled subset of <see cref="AllWriteTools"/>. Both
        /// <see cref="OutlookToolHost"/> (tool registration) and
        /// <see cref="Services.Tools.ToolCatalogSchema"/> (request-time
        /// catalog) consult this set when WriteToolsEnabled=true.
        /// </summary>
        public static HashSet<string> EnabledWriteTools { get; set; } =
            new HashSet<string>(AllWriteTools, StringComparer.Ordinal);

        /// <summary>
        /// Raised on the UI thread after Settings saves AI settings or refreshes
        /// the model catalog, so open panes re-read their reasoning options.
        /// </summary>
        public static event EventHandler AiSettingsChanged;

        public static void NotifyAiSettingsChanged()
        {
            var handlers = AiSettingsChanged;
            if (handlers == null) return;
            // One pane failing (e.g. mid-dispose) must not starve the others.
            foreach (EventHandler handler in handlers.GetInvocationList())
            {
                try
                {
                    handler(null, EventArgs.Empty);
                }
                catch (Exception ex)
                {
                    OutlookAI.Diagnostics.TraceLog.Write("AiSettingsChanged handler failed: " + ex.Message, "Config");
                }
            }
        }

        /// <summary>
        /// Reasoning-effort options for <paramref name="model"/> from the
        /// catalog: <c>Auto</c> first, then the efforts that model accepts
        /// (e.g. gpt-5.5 has neither Minimal nor Max). The wire value comes
        /// from <see cref="ModelCatalog.ResolveWireEffort"/>.
        /// </summary>
        public static string[] ReasoningEffortsForModel(string model)
        {
            return ModelCatalog.EffortsFor(model);
        }

        // ============================================================
        // END CONFIGURATION
        // ============================================================

        // Global config: admin-controlled, applies to all users on this server
        private static readonly string GlobalConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "OutlookAI",
            "config.xml"
        );

        // Per-user config: may override AdminPassword only
        private static readonly string UserConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OutlookAI",
            "config.xml"
        );

        // Shared admin defaults: writable by Admins (RDS scenario), readable by
        // all users on the box. Loaded between the server-authoritative global
        // config and the per-user AppData override. SettingsForm-driven Save
        // writes here AND to AppData so the admin's preference becomes the
        // default for every user on the server.
        private static readonly string SharedConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "OutlookAI",
            "config.xml"
        );

        static Config()
        {
            LoadConfig();
        }

        public static void LoadConfig()
        {
            ModelCatalog = LoadModelCatalog();
            ReloadConfigFiles();
        }

        /// <summary>
        /// Re-applies the three config.xml layers against the current
        /// <see cref="ModelCatalog"/>. Settings calls it after Update Models so
        /// values the previous catalog rejected (e.g. a new model named in
        /// config.xml) take effect instead of being overwritten by the next Save.
        /// </summary>
        public static void ReloadConfigFiles()
        {
            LoadConfigFromPaths(GlobalConfigFilePath, SharedConfigFilePath, UserConfigFilePath);
        }

        // The cached models.json, unless this build's own list is newer (e.g.
        // an update shipped retirement dates the cache predates).
        private static ModelCatalog LoadModelCatalog()
        {
            ModelCatalog cached = null;
            try
            {
                cached = new ModelCatalogStore().Load();
            }
            catch (Exception ex)
            {
                OutlookAI.Diagnostics.TraceLog.Write("Model catalog load failed: " + ex.Message, "Config");
            }
            var chosen = ModelCatalog.Newest(cached, BuiltInModelCatalog.Instance);
            OutlookAI.Diagnostics.TraceLog.Write(
                "Model catalog: " + (chosen.IsBuiltIn ? "built-in" : "models.json") + ", "
                + chosen.ListedSlugs.Count + " listed models as of "
                + (chosen.FetchedAt.HasValue ? chosen.FetchedAt.Value.ToString("o") : "?")
                + " (client_version " + chosen.ClientVersion + ")",
                "Config");
            return chosen;
        }

        // Test seam: explicit paths so we don't touch Program Files /
        // ProgramData / AppData during unit tests. The layers merge into a
        // scratch copy that is published at the end, so a request built while
        // Update Models reloads the files never sees transient defaults.
        public static void LoadConfigFromPaths(string globalConfigPath, string sharedConfigPath, string userConfigPath)
        {
            var values = Values.Defaults();
            LoadFromFile(globalConfigPath, values, allowServerFields: true);
            LoadFromFile(sharedConfigPath, values, allowServerFields: false);
            LoadFromFile(userConfigPath, values, allowServerFields: false);
            Apply(values);
        }

        // Back-compat overload for existing tests that don't care about the
        // shared-defaults layer. Equivalent to passing a non-existent shared
        // path (LoadFromFile short-circuits on null/missing).
        public static void LoadConfigFromPaths(string globalConfigPath, string userConfigPath)
        {
            LoadConfigFromPaths(globalConfigPath, sharedConfigPath: null, userConfigPath);
        }

        public static void ResetDefaults()
        {
            Apply(Values.Defaults());
        }

        /// <summary>
        /// Every &lt;Model&gt; the config layers named at the last load, in load
        /// order, including ones the catalog rejected. Update Models keeps these
        /// alive when a filtered model list leaves them out.
        /// </summary>
        public static IReadOnlyList<string> ModelsNamedInConfig { get; private set; } = new string[0];

        // Scratch copy of the loadable settings (see LoadConfigFromPaths).
        private sealed class Values
        {
            public string AdminPassword;
            public string CodexAuthPath;
            public string Model;
            public string VoiceModel;
            public string ModelCatalogClientVersion;
            public string ReasoningEffort;
            public bool WriteToolsEnabled;
            public int MaxBulkExportRows;
            public HashSet<string> EnabledWriteTools;
            public readonly List<string> ModelsNamed = new List<string>();

            public static Values Defaults()
            {
                return new Values
                {
                    AdminPassword = "admin",
                    CodexAuthPath = DefaultCodexAuthPath,
                    Model = DefaultModel,
                    VoiceModel = DefaultVoiceModel,
                    ModelCatalogClientVersion = "",
                    ReasoningEffort = DefaultReasoningEffort,
                    WriteToolsEnabled = DefaultWriteToolsEnabled,
                    MaxBulkExportRows = DefaultMaxBulkExportRows,
                    EnabledWriteTools = new HashSet<string>(AllWriteTools, StringComparer.Ordinal),
                };
            }
        }

        private static void Apply(Values values)
        {
            AdminPassword = values.AdminPassword;
            CodexAuthPath = values.CodexAuthPath;
            Model = values.Model;
            VoiceModel = values.VoiceModel;
            ModelCatalogClientVersion = values.ModelCatalogClientVersion;
            ReasoningEffort = values.ReasoningEffort;
            WriteToolsEnabled = values.WriteToolsEnabled;
            MaxBulkExportRows = values.MaxBulkExportRows;
            EnabledWriteTools = values.EnabledWriteTools;
            ModelsNamedInConfig = values.ModelsNamed.AsReadOnly();
        }

        private static void LoadFromFile(string filePath, Values values, bool allowServerFields)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    return;
                }

                var doc = XDocument.Load(filePath);
                var root = doc.Root;
                if (root == null)
                {
                    return;
                }

                // User-tunable fields (AdminPassword, ReasoningEffort,
                // WriteToolsEnabled) are read from both global and per-user
                // config. Per-user takes precedence because it's loaded
                // second. Settings UI persists them via SaveConfig.
                var adminPassword = root.Element("AdminPassword");
                if (adminPassword != null && !string.IsNullOrEmpty(adminPassword.Value))
                {
                    values.AdminPassword = adminPassword.Value;
                }

                // "Auto" (or its old name "None") or any effort some catalog
                // model accepts (e.g. "Max"), stored in canonical casing;
                // anything else keeps the prior value.
                var reasoningEffort = root.Element("ReasoningEffort");
                if (reasoningEffort != null && !string.IsNullOrWhiteSpace(reasoningEffort.Value))
                {
                    var normalized = ModelCatalog.NormalizeEffort(reasoningEffort.Value);
                    if (normalized != null)
                    {
                        values.ReasoningEffort = normalized;
                    }
                    else
                    {
                        TraceIgnored("ReasoningEffort", reasoningEffort.Value, filePath);
                    }
                }

                var writeToolsEnabled = root.Element("WriteToolsEnabled");
                if (writeToolsEnabled != null && bool.TryParse(writeToolsEnabled.Value, out var wte))
                {
                    values.WriteToolsEnabled = wte;
                }

                var enabledWriteTools = root.Element("EnabledWriteTools");
                if (enabledWriteTools != null && !string.IsNullOrWhiteSpace(enabledWriteTools.Value))
                {
                    // Comma-separated list; intersect with the canonical set
                    // so unknown tool names (typo / future tool removed) are
                    // silently dropped instead of breaking the dispatcher.
                    var requested = enabledWriteTools.Value
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(x => x.Trim());
                    var canonical = new HashSet<string>(AllWriteTools, StringComparer.Ordinal);
                    values.EnabledWriteTools = new HashSet<string>(
                        requested.Where(canonical.Contains),
                        StringComparer.Ordinal);
                }

                // Model is user-tunable (Settings UI). Server provides the
                // default, but per-user override beats it on load. Names the
                // model catalog doesn't know (typos, Claude-era v1 values,
                // retired models) fall back to whatever was already set;
                // hidden catalog models are accepted when set explicitly.
                var model = root.Element("Model");
                if (model != null && !string.IsNullOrWhiteSpace(model.Value))
                {
                    values.ModelsNamed.Add(model.Value.Trim());
                    var entry = ModelCatalog.Find(model.Value);
                    if (entry != null)
                    {
                        values.Model = entry.Slug;
                    }
                    else
                    {
                        TraceIgnored("Model", model.Value, filePath);
                    }
                }

                if (!allowServerFields)
                {
                    return;
                }

                var codexAuthPath = root.Element("CodexAuthPath");
                if (codexAuthPath != null && !string.IsNullOrWhiteSpace(codexAuthPath.Value))
                {
                    values.CodexAuthPath = codexAuthPath.Value;
                }

                var voiceModel = root.Element("VoiceModel");
                if (voiceModel != null && !string.IsNullOrWhiteSpace(voiceModel.Value))
                {
                    values.VoiceModel = voiceModel.Value;
                }

                var catalogClientVersion = root.Element("ModelCatalogClientVersion");
                if (catalogClientVersion != null && !string.IsNullOrWhiteSpace(catalogClientVersion.Value))
                {
                    values.ModelCatalogClientVersion = catalogClientVersion.Value.Trim();
                }

                var maxBulkExportRows = root.Element("MaxBulkExportRows");
                if (maxBulkExportRows != null && int.TryParse(maxBulkExportRows.Value, out var mber))
                {
                    if (mber < MinBulkExportRows) mber = MinBulkExportRows;
                    if (mber > MaxBulkExportRowsCeiling) mber = MaxBulkExportRowsCeiling;
                    values.MaxBulkExportRows = mber;
                }
            }
            catch
            {
                // Skip if file is missing or invalid; defaults stay in place.
            }
        }

        // Values from a config.xml the model catalog doesn't know are dropped;
        // leave a trail so a "my setting didn't stick" report is diagnosable.
        private static void TraceIgnored(string element, string value, string filePath)
        {
            OutlookAI.Diagnostics.TraceLog.Write(
                "Config: ignoring <" + element + ">" + value.Trim() + "</" + element + "> in " + filePath
                + " (not offered by the " + (ModelCatalog.IsBuiltIn ? "built-in" : "cached") + " model list)",
                "Config");
        }

        public static void SaveConfig()
        {
            // Per-user config persists AdminPassword + the user-tunable AI
            // behavior fields. Shared config persists the same fields so
            // admins on an RDS host can set server-wide defaults that
            // propagate to every user.
            var doc = BuildSavedConfig();

            // 1) Per-user override (always attempted; failures silently
            //    swallowed so a read-only AppData doesn't block the workflow).
            TrySaveTo(UserConfigFilePath, doc);

            // 2) Shared admin defaults (only succeeds if running as
            //    Administrator on the RDS host - ProgramData ACL by default
            //    is Admin-write, User-read. Regular users silently no-op
            //    here, which is the intended behavior: a non-admin can't
            //    change server-wide defaults).
            TrySaveTo(SharedConfigFilePath, doc);
        }

        internal static XDocument BuildSavedConfig()
        {
            return new XDocument(
                new XElement("Config",
                    new XElement("AdminPassword", AdminPassword),
                    new XElement("Model", Model),
                    new XElement("ReasoningEffort", ReasoningEffortNames.ToConfigValue(ReasoningEffort)),
                    new XElement("WriteToolsEnabled", WriteToolsEnabled),
                    new XElement("EnabledWriteTools",
                        string.Join(",", EnabledWriteTools ?? new HashSet<string>()))
                )
            );
        }

        private static void TrySaveTo(string filePath, XDocument doc)
        {
            try
            {
                var dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                doc.Save(filePath);
            }
            catch (Exception ex)
            {
                // Silently fail (read-only target, ACL denied, etc.) but record
                // it so the maintainer can see the failure mode without breaking
                // the user flow. Without this trace the admin gets a "Saved"
                // indicator and only discovers the shared write didn't take when
                // another user's login still shows defaults.
                try
                {
                    OutlookAI.Diagnostics.TraceLog.Write(
                        "Config.TrySaveTo failed for '" + filePath + "': " + ex.Message,
                        "Config");
                }
                catch { }
            }
        }
    }
}
