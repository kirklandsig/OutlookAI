using System;

namespace OutlookAI.Services.Models
{
    /// <summary>
    /// Fallback catalog used until the first Settings → Update Models, and
    /// whenever no cached models.json is readable. Mirrors the listed models from
    /// a live /backend-api/codex/models probe (client_version 0.155.1,
    /// 2026-09-22). Efforts are kept exactly as the catalog lists them, Codex's
    /// client-only "ultra" included; <see cref="ModelCatalog"/> filters them.
    /// </summary>
    public static class BuiltInModelCatalog
    {
        public static readonly DateTimeOffset AsOf = new DateTimeOffset(2026, 9, 22, 0, 0, 0, TimeSpan.Zero);

        private static readonly string[] ThroughUltra = { "low", "medium", "high", "xhigh", "max", "ultra" };
        private static readonly string[] ThroughMax = { "low", "medium", "high", "xhigh", "max" };
        private static readonly string[] ThroughXHigh = { "low", "medium", "high", "xhigh" };

        public static readonly ModelCatalog Instance = new ModelCatalog(
            new[]
            {
                Listed("gpt-6-astra", "GPT-6-Astra", "Our most capable model for complex, demanding work.", 1, ThroughUltra, null),
                Listed("gpt-6-sol", "GPT-6-Sol", "GPT-6 Sol Codex model.", 2, ThroughUltra, null),
                Listed("gpt-6-luna", "GPT-6-Luna", "GPT-6 Luna Codex model.", 3, ThroughMax, null),
                Listed("gpt-5.6-sol", "GPT-5.6-Sol", "Reliable agentic workhorse for everyday tasks.", 4, ThroughUltra,
                    new ModelUpgradeInfo("gpt-6-sol", null)),
                Listed("gpt-5.6-terra", "GPT-5.6-Terra", "Balanced agentic coding model for everyday work.", 7, ThroughUltra,
                    new ModelUpgradeInfo("gpt-6-sol", null)),
                Listed("gpt-5.6-luna", "GPT-5.6-Luna", "Fast and affordable agentic coding model.", 8, ThroughMax,
                    new ModelUpgradeInfo("gpt-6-luna", null)),
                Listed("gpt-5.5", "GPT-5.5", "Proven previous-generation model for coding and general work.", 12, ThroughXHigh,
                    new ModelUpgradeInfo("gpt-5.6-sol", new DateTimeOffset(2026, 10, 14, 19, 0, 0, TimeSpan.Zero))),
            },
            serverEfforts: null,
            source: ModelCatalog.SourceBuiltIn,
            fetchedAt: AsOf,
            clientVersion: CodexClientVersion.BuiltInFloor,
            clientVersionSource: "built-in");

        private static ModelCatalogEntry Listed(
            string slug, string displayName, string description, int priority, string[] efforts, ModelUpgradeInfo upgrade)
        {
            return new ModelCatalogEntry(slug, displayName, description, true, priority, efforts, upgrade);
        }
    }
}
