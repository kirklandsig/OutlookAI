using System;
using System.Collections.Generic;
using OutlookAI.Services.Models;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Pins the Config statics that model resolution reads (catalog, clock) and
    /// restores them plus the settings Settings saves on Dispose. Without it, a test
    /// would see whatever models.json / config.xml the dev box has, and anything
    /// touching gpt-5.5 would change behavior on its 2026-10-14 retirement.
    /// Use only from [Collection("Config")] classes.
    /// </summary>
    internal sealed class ConfigStateScope : IDisposable
    {
        private readonly ModelCatalog _catalog;
        private readonly Func<DateTimeOffset> _clock;
        private readonly string _model;
        private readonly string _effort;
        private readonly string _adminPassword;
        private readonly bool _writeToolsEnabled;
        private readonly HashSet<string> _enabledWriteTools;

        public ConfigStateScope(ModelCatalog catalog = null, DateTimeOffset? now = null)
        {
            _catalog = Config.ModelCatalog;
            _clock = Config.Clock;
            _model = Config.Model;
            _effort = Config.ReasoningEffort;
            _adminPassword = Config.AdminPassword;
            _writeToolsEnabled = Config.WriteToolsEnabled;
            _enabledWriteTools = Config.EnabledWriteTools;

            Config.ModelCatalog = catalog ?? BuiltInModelCatalog.Instance;
            var fixedNow = now ?? TestCatalogs.FixedNow;
            Config.Clock = () => fixedNow;
        }

        public void Dispose()
        {
            Config.ModelCatalog = _catalog;
            Config.Clock = _clock;
            Config.Model = _model;
            Config.ReasoningEffort = _effort;
            Config.AdminPassword = _adminPassword;
            Config.WriteToolsEnabled = _writeToolsEnabled;
            Config.EnabledWriteTools = _enabledWriteTools;
        }
    }
}
