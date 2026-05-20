using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

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

        public const string DefaultModel = "gpt-5.5";
        public const string DefaultVoiceModel = "gpt-realtime-1.5";
        public const string DefaultCodexAuthPath = @"C:\ProgramData\OutlookAI\auth.json";
        public const string DefaultReasoningEffort = "None";
        public const bool DefaultWriteToolsEnabled = true;

        public static string AdminPassword { get; set; } = "admin";
        public static string CodexAuthPath { get; set; } = DefaultCodexAuthPath;
        public static string Model { get; set; } = DefaultModel;
        public static string VoiceModel { get; set; } = DefaultVoiceModel;

        /// <summary>
        /// Default reasoning effort sent to the Codex backend on each turn.
        /// One of <see cref="AvailableReasoningEfforts"/>. "None" means omit the
        /// reasoning block entirely. Per-turn overrides via
        /// <c>ConversationContext.ReasoningEffortOverride</c>.
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

        public static readonly string[] AvailableModels =
        {
            "gpt-5.5",
            "gpt-5.5-pro",
            "gpt-5.4",
            "gpt-5.4-mini",
            "gpt-4.1-mini",
            "gpt-4.1-nano",
            "gpt-5.3-codex"
        };

        // Authoritative source: OpenAI docs (Reasoning models guide) +
        // confirmed by Codex backend error messages on production traffic:
        //   "Supported values are model-dependent and can include
        //    'none', 'minimal', 'low', 'medium', 'high', and 'xhigh'."
        public static readonly string[] AvailableReasoningEfforts =
        {
            "None",
            "Minimal",
            "Low",
            "Medium",
            "High",
            "XHigh"
        };

        /// <summary>
        /// Returns the reasoning-effort options that are valid for a given
        /// model. Per-model overrides because not every model supports every
        /// value:
        ///   - gpt-5.5 family rejects 'minimal' (confirmed via backend error).
        ///   - gpt-4.1-mini / nano are non-reasoning - only 'none' is valid.
        ///   - gpt-5.4 family supports the full set including 'minimal'.
        /// Wire format is the lowercased value; see
        /// <c>CodexChatService.BuildRunTurnRequest</c>.
        /// </summary>
        public static string[] ReasoningEffortsForModel(string model)
        {
            if (model == "gpt-4.1-mini" || model == "gpt-4.1-nano")
            {
                return new[] { "None" };
            }
            if (model == "gpt-5.5" || model == "gpt-5.5-pro" || model == "gpt-5.3-codex")
            {
                // 'Minimal' is unsupported on gpt-5.5 per the backend error
                // message: "'minimal' is not supported with the 'gpt-5.5' model.
                // Supported values are: 'none', 'low', 'medium', 'high', and
                // 'xhigh'."
                return new[] { "None", "Low", "Medium", "High", "XHigh" };
            }
            // gpt-5.4, gpt-5.4-mini, and any future model default to the
            // full enum.
            return AvailableReasoningEfforts;
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

        static Config()
        {
            LoadConfig();
        }

        public static void LoadConfig()
        {
            LoadConfigFromPaths(GlobalConfigFilePath, UserConfigFilePath);
        }

        // Test seam: explicit paths so we don't touch Program Files / AppData
        // during unit tests.
        public static void LoadConfigFromPaths(string globalConfigPath, string userConfigPath)
        {
            ResetDefaults();
            LoadFromFile(globalConfigPath, allowServerFields: true);
            LoadFromFile(userConfigPath, allowServerFields: false);
        }

        public static void ResetDefaults()
        {
            AdminPassword = "admin";
            CodexAuthPath = DefaultCodexAuthPath;
            Model = DefaultModel;
            VoiceModel = DefaultVoiceModel;
            ReasoningEffort = DefaultReasoningEffort;
            WriteToolsEnabled = DefaultWriteToolsEnabled;
            EnabledWriteTools = new HashSet<string>(AllWriteTools, StringComparer.Ordinal);
        }

        private static void LoadFromFile(string filePath, bool allowServerFields)
        {
            try
            {
                if (!File.Exists(filePath))
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
                    AdminPassword = adminPassword.Value;
                }

                var reasoningEffort = root.Element("ReasoningEffort");
                if (reasoningEffort != null && !string.IsNullOrWhiteSpace(reasoningEffort.Value))
                {
                    foreach (var allowed in AvailableReasoningEfforts)
                    {
                        if (string.Equals(allowed, reasoningEffort.Value, StringComparison.OrdinalIgnoreCase))
                        {
                            ReasoningEffort = allowed;
                            break;
                        }
                    }
                }

                var writeToolsEnabled = root.Element("WriteToolsEnabled");
                if (writeToolsEnabled != null && bool.TryParse(writeToolsEnabled.Value, out var wte))
                {
                    WriteToolsEnabled = wte;
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
                    EnabledWriteTools = new HashSet<string>(
                        requested.Where(canonical.Contains),
                        StringComparer.Ordinal);
                }

                // Model is user-tunable in Phase 2 (Settings UI lets the admin
                // pick from AvailableModels). Server provides the default, but
                // per-user override beats it on load. Unknown model names fall
                // back to whatever was already set.
                var model = root.Element("Model");
                if (model != null && !string.IsNullOrWhiteSpace(model.Value)
                    && AvailableModels.Contains(model.Value))
                {
                    Model = model.Value;
                }

                if (!allowServerFields)
                {
                    return;
                }

                var codexAuthPath = root.Element("CodexAuthPath");
                if (codexAuthPath != null && !string.IsNullOrWhiteSpace(codexAuthPath.Value))
                {
                    CodexAuthPath = codexAuthPath.Value;
                }

                var voiceModel = root.Element("VoiceModel");
                if (voiceModel != null && !string.IsNullOrWhiteSpace(voiceModel.Value))
                {
                    VoiceModel = voiceModel.Value;
                }
            }
            catch
            {
                // Skip if file is missing or invalid; defaults stay in place.
            }
        }

        public static void SaveConfig()
        {
            try
            {
                var dir = Path.GetDirectoryName(UserConfigFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Per-user config persists AdminPassword + the user-tunable
                // AI behavior fields (ReasoningEffort, WriteToolsEnabled,
                // EnabledWriteTools, Model). All other fields are
                // server-authoritative and owned by the global config in
                // Program Files.
                var doc = new XDocument(
                    new XElement("Config",
                        new XElement("AdminPassword", AdminPassword),
                        new XElement("Model", Model),
                        new XElement("ReasoningEffort", ReasoningEffort),
                        new XElement("WriteToolsEnabled", WriteToolsEnabled),
                        new XElement("EnabledWriteTools",
                            string.Join(",", EnabledWriteTools ?? new HashSet<string>()))
                    )
                );

                doc.Save(UserConfigFilePath);
            }
            catch
            {
                // Silently fail if can't save (e.g. read-only AppData).
            }
        }
    }
}
