using System;
using OutlookAI.Services.Models;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Pins the Config statics that model resolution reads (catalog, clock) and
    /// restores them plus Model / ReasoningEffort on Dispose. Without it, a test
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

        public ConfigStateScope(ModelCatalog catalog = null, DateTimeOffset? now = null)
        {
            _catalog = Config.ModelCatalog;
            _clock = Config.Clock;
            _model = Config.Model;
            _effort = Config.ReasoningEffort;

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
        }
    }
}
