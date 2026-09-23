using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using OutlookAI.Services.Models;

namespace OutlookAI
{
    public static class Config
    {
        // ============================================================
        // CONFIGURATION DEFAULTS (v2 - ChatGPT OAuth)
        // Loaded as defaults -> global config (Program Files; the only
        // source of server fields such as CodexAuthPath) -> a per-user
        // AppData config -> the settings Settings saves for every user
        // (ProgramData), later layers winning setting by setting. Legacy v1
        // elements (ApiKey, OpenAIApiKey, WhisperModel, TranscribeModel,
        // MaxTokens, and Claude model names) are ignored if encountered.
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

        // Global config: admin-controlled, applies to all users on this server.
        // %ProgramW6432% so a 32-bit Outlook still finds the installer's
        // C:\Program Files\OutlookAI, not Program Files (x86).
        private static readonly string GlobalConfigFilePath = Path.Combine(
            Environment.GetEnvironmentVariable("ProgramW6432")
                ?? Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "OutlookAI",
            "config.xml"
        );

        // Per-user config, which Settings also wrote before v2.2.2. Loaded
        // before SharedConfigFilePath, so it only fills in what that doesn't set.
        private static readonly string UserConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OutlookAI",
            "config.xml"
        );

        // Server-wide settings: what Settings saves, for every user on the
        // machine (the installer grants Authenticated Users Modify on this
        // folder). Loaded last.
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
            LoadConfigFromPaths(GlobalConfigFilePath, SharedConfigFilePath, UserConfigFilePath);
        }

        /// <summary>
        /// Re-applies the three config.xml layers against the current
        /// <see cref="ModelCatalog"/>. Called when Settings opens (so it starts
        /// from what is saved now), after a failed save, and after Update Models
        /// (values the previous catalog rejected, e.g. a new model named in
        /// config.xml, may be valid now).
        /// </summary>
        public static void ReloadConfigFiles()
        {
            ReloadConfigFilesFrom(GlobalConfigFilePath, SharedConfigFilePath, UserConfigFilePath);
        }

        // Test seam for ReloadConfigFiles. Unlike loading at startup, it keeps
        // the settings this session has when a config.xml that exists can't be
        // read (e.g. another program holds it): what the other layers say isn't
        // what was saved.
        internal static void ReloadConfigFilesFrom(string globalConfigPath, string sharedConfigPath, string userConfigPath)
        {
            bool complete;
            var values = ReadLayers(globalConfigPath, sharedConfigPath, userConfigPath, out complete);
            if (complete)
            {
                Apply(values);
            }
            else
            {
                OutlookAI.Diagnostics.TraceLog.Write("Config: keeping the current settings; a config file could not be read", "Config");
            }
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
            bool complete;
            Apply(ReadLayers(globalConfigPath, sharedConfigPath, userConfigPath, out complete));
        }

        // Settings saves for every user (the shared file), so that is read last
        // and wins. A per-user file, which Settings also wrote before v2.2.2,
        // only fills in what the shared file doesn't set. complete is false when
        // a file that exists couldn't be read.
        private static Values ReadLayers(string globalConfigPath, string sharedConfigPath, string userConfigPath, out bool complete)
        {
            var values = Values.Defaults();
            complete = LoadFromFile(globalConfigPath, values, allowServerFields: true);
            complete &= LoadFromFile(userConfigPath, values, allowServerFields: false);
            complete &= LoadFromFile(sharedConfigPath, values, allowServerFields: false);
            return values;
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

        // False when the file may exist but couldn't be read. Only "not found"
        // counts as no file: File.Exists also says false when access is denied.
        private static bool LoadFromFile(string filePath, Values values, bool allowServerFields)
        {
            if (string.IsNullOrEmpty(filePath)) return true;
            try
            {
                XDocument doc;
                try
                {
                    doc = LoadXmlWhenFree(filePath);
                }
                catch (Exception ex) when (ex is FileNotFoundException || ex is DirectoryNotFoundException)
                {
                    return true;
                }
                var root = doc.Root;
                if (root == null)
                {
                    return true;
                }

                // User-tunable fields (AdminPassword, ReasoningEffort, write
                // tools, Model) are read from every layer; a later layer wins.
                // Settings saves them via SaveSettings.
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
                if (enabledWriteTools != null)
                {
                    values.EnabledWriteTools = ParseWriteTools(enabledWriteTools.Value);
                }

                // Model is set in Settings; a later layer wins (see
                // LoadConfigFromPaths). Names the model catalog doesn't know
                // (typos, Claude-era v1 values, retired models) fall back to
                // whatever was already set; hidden catalog models are accepted
                // when set explicitly.
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
                    return true;
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
                return true;
            }
            catch (Exception ex)
            {
                // Skip if the file is invalid or can't be opened; the other
                // layers stay in place.
                OutlookAI.Diagnostics.TraceLog.Write("Config: could not read " + filePath + ": " + ex.Message, "Config");
                return false;
            }
        }

        // A Settings save in another session swaps config.xml in, and a read
        // that lands in that moment fails; try again briefly rather than run
        // without the file until Outlook restarts.
        private static XDocument LoadXmlWhenFree(string path)
        {
            for (var attempt = 1; ; attempt++)
            {
                try
                {
                    return XDocument.Load(path);
                }
                catch (IOException ex) when (attempt < 20 && IsBusy(ex))
                {
                    Thread.Sleep(50);
                }
            }
        }

        // Another process has the file open (sharing or lock violation).
        private static bool IsBusy(IOException ex)
        {
            var code = ex.HResult & 0xFFFF;
            return code == 32 || code == 33;
        }

        // A saved EnabledWriteTools list: comma-separated (empty = every tool
        // unchecked), intersected with the canonical set so unknown tool names
        // (typo / future tool removed) are silently dropped instead of breaking
        // the dispatcher.
        private static HashSet<string> ParseWriteTools(string list)
        {
            var canonical = new HashSet<string>(AllWriteTools, StringComparer.Ordinal);
            return new HashSet<string>(
                list.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(canonical.Contains),
                StringComparer.Ordinal);
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

        /// <summary>The settings Settings saves, as named in config.xml.</summary>
        internal static readonly string[] SavedSettingNames =
            { "AdminPassword", "Model", "ReasoningEffort", "WriteToolsEnabled", "EnabledWriteTools" };

        private static readonly TimeSpan SaveLockTimeout = TimeSpan.FromSeconds(5);

        /// <summary>
        /// Saves settings for every user on this machine, in the server-wide
        /// config.xml (ProgramData): the <paramref name="changed"/> ones with this
        /// session's values, the others as already saved there (another admin
        /// may have changed them since this Outlook loaded them). Given
        /// <paramref name="toolsBefore"/>, the write tools this session's change
        /// started from, the tools it switched on or off are applied one by one
        /// to the saved list instead, and this session takes the result.
        /// Returns null once saved, otherwise why it wasn't.
        /// </summary>
        public static string SaveSettings(string[] changed, ISet<string> toolsBefore = null)
        {
            return SaveSettingsTo(SharedConfigFilePath, changed, toolsBefore);
        }

        // Test seam: explicit path and lock timeout.
        internal static string SaveSettingsTo(
            string sharedConfigPath, IEnumerable<string> changed, ISet<string> toolsBefore = null, TimeSpan? lockTimeout = null)
        {
            var temp = sharedConfigPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                var dir = Path.GetDirectoryName(sharedConfigPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                // Read, merge and replace under a cross-process lock, so two
                // sessions saving at once can't undo each other's changes.
                using (Services.FileLock.Acquire(sharedConfigPath + ".lock", lockTimeout ?? SaveLockTimeout))
                {
                    var onDisk = LoadForSave(sharedConfigPath);
                    var changedSet = new HashSet<string>(changed ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
                    var tools = toolsBefore != null && changedSet.Contains("EnabledWriteTools")
                        ? MergeWriteTools(onDisk, toolsBefore)
                        : null;
                    var doc = MergeSettings(onDisk, changedSet, tools);
                    // Written beside config.xml, flushed and checked, then swapped
                    // in whole, so a failed or interrupted save never leaves a
                    // truncated file behind.
                    using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        doc.Save(stream);
                        stream.Flush(true);
                    }
                    XDocument.Load(temp);
                    for (var attempt = 1; ; attempt++)
                    {
                        try
                        {
                            if (File.Exists(sharedConfigPath))
                            {
                                File.Replace(temp, sharedConfigPath, null);
                            }
                            else
                            {
                                File.Move(temp, sharedConfigPath);
                            }
                            break;
                        }
                        catch (Exception) when (!File.Exists(sharedConfigPath) && !Directory.Exists(sharedConfigPath) && File.Exists(temp))
                        {
                            // File.Replace can fail after config.xml is already gone;
                            // the new file is complete and checked, so finish the swap.
                            File.Move(temp, sharedConfigPath);
                            break;
                        }
                        catch (IOException ex) when (attempt < 20 && IsBusy(ex))
                        {
                            // Another session is reading config.xml for a moment.
                            Thread.Sleep(50);
                        }
                    }
                    if (tools != null)
                    {
                        EnabledWriteTools = tools;
                        WriteToolsEnabled = tools.Count > 0;
                    }
                }
                return null;
            }
            catch (TimeoutException)
            {
                TraceSave(sharedConfigPath, "timed out waiting for another session's save");
                return "Another Outlook session is saving the settings. Try again.";
            }
            catch (Exception ex)
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
                TraceSave(sharedConfigPath, ex.Message);
                return ex.Message;
            }
        }

        // What is saved on disk, with this session's values for the changed
        // settings (and the merged write tools, when given). Settings the file
        // doesn't set (missing, or blank where the loader skips blanks) are
        // added from this session too, so it always holds all of them.
        private static XDocument MergeSettings(XDocument onDisk, ISet<string> changed, HashSet<string> tools)
        {
            foreach (var name in changed)
            {
                if (Array.IndexOf(SavedSettingNames, name) < 0) throw new ArgumentException("Not a saved setting: " + name);
            }
            var doc = onDisk != null && onDisk.Root != null ? onDisk : new XDocument(new XElement("Config"));
            foreach (var name in SavedSettingNames)
            {
                var element = tools == null ? SettingElement(name)
                    : name == "EnabledWriteTools" ? new XElement(name, string.Join(",", tools))
                    : name == "WriteToolsEnabled" ? new XElement(name, tools.Count > 0)
                    : SettingElement(name);
                var existing = doc.Root.Element(name);
                // A blank tool list is a real value: every tool unchecked.
                var unset = existing == null || (name != "EnabledWriteTools" && string.IsNullOrWhiteSpace(existing.Value));
                if (existing == null)
                {
                    doc.Root.Add(element);
                }
                else if (unset || changed.Contains(name))
                {
                    existing.ReplaceWith(element);
                }
            }
            return doc;
        }

        // The saved write tools with this session's changes since toolsBefore
        // applied one by one, so another admin's changes to other tools stay.
        // When the saved file doesn't settle the tools by itself, the changes
        // apply to toolsBefore (what this session had from all the layers).
        private static HashSet<string> MergeWriteTools(XDocument onDisk, ISet<string> toolsBefore)
        {
            var mine = WriteToolsEnabled
                ? new HashSet<string>(EnabledWriteTools ?? new HashSet<string>(), StringComparer.Ordinal)
                : new HashSet<string>(StringComparer.Ordinal);
            var tools = SavedWriteToolsIn(onDisk) ?? new HashSet<string>(toolsBefore, StringComparer.Ordinal);
            tools.UnionWith(mine.Where(t => !toolsBefore.Contains(t)));
            tools.ExceptWith(toolsBefore.Where(t => !mine.Contains(t)));
            return tools;
        }

        // The write tools in effect per a saved file: none while its switch is
        // off, else its list. Null when it lacks the switch, or the list while
        // on: the loader takes those from the other layers.
        private static HashSet<string> SavedWriteToolsIn(XDocument doc)
        {
            var root = doc == null ? null : doc.Root;
            var enabled = root == null ? null : root.Element("WriteToolsEnabled");
            if (enabled == null || !bool.TryParse(enabled.Value, out var on)) return null;
            if (!on) return new HashSet<string>(StringComparer.Ordinal);
            var list = root.Element("EnabledWriteTools");
            return list == null ? null : ParseWriteTools(list.Value);
        }

        // What is saved there now, or null when there is no file yet (the save
        // then writes a complete one). A file that can't be read stops the save
        // rather than be replaced with this session's values, which may come
        // from an old per-user copy.
        private static XDocument LoadForSave(string path)
        {
            try
            {
                return LoadXmlWhenFree(path);
            }
            catch (FileNotFoundException)
            {
                return null;
            }
            catch (XmlException ex)
            {
                throw new InvalidOperationException(
                    path + " is not valid XML (" + ex.Message + "). Fix or delete it, then save again.", ex);
            }
        }

        // One saved setting as config.xml stores it.
        private static XElement SettingElement(string name)
        {
            switch (name)
            {
                case "AdminPassword": return new XElement(name, AdminPassword);
                case "Model": return new XElement(name, Model);
                case "ReasoningEffort": return new XElement(name, ReasoningEffortNames.ToConfigValue(ReasoningEffort));
                case "WriteToolsEnabled": return new XElement(name, WriteToolsEnabled);
                case "EnabledWriteTools": return new XElement(name, string.Join(",", EnabledWriteTools ?? new HashSet<string>()));
                default: throw new ArgumentException("Not a saved setting: " + name);
            }
        }

        internal static XDocument BuildSavedConfig()
        {
            return new XDocument(new XElement("Config", SavedSettingNames.Select(SettingElement)));
        }

        private static void TraceSave(string filePath, string message)
        {
            try
            {
                OutlookAI.Diagnostics.TraceLog.Write("Config: saving '" + filePath + "': " + message, "Config");
            }
            catch { }
        }
    }
}
