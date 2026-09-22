using System;
using OutlookAI.Services.Models;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Fixture catalogs for model-catalog tests. <see cref="FixedNow"/> sits
    /// before the built-in gpt-5.5 retirement date (2026-10-14) so tests that
    /// use the built-in catalog never change behavior with the wall clock.
    /// </summary>
    internal static class TestCatalogs
    {
        public static readonly DateTimeOffset FixedNow =
            new DateTimeOffset(2026, 9, 22, 12, 0, 0, TimeSpan.Zero);

        public static ModelCatalogEntry Entry(
            string slug,
            int priority,
            string[] efforts,
            bool listed = true,
            ModelUpgradeInfo upgrade = null)
        {
            return new ModelCatalogEntry(
                slug, slug.ToUpperInvariant(), "About " + slug, listed, priority, efforts, upgrade);
        }

        public static ModelCatalog Catalog(params ModelCatalogEntry[] entries)
        {
            return new ModelCatalog(entries, null, ModelCatalog.SourceChatGpt, FixedNow, "0.155.1");
        }

        /// <summary>One listed model that accepts every effort the wire tests use.</summary>
        public static ModelCatalog AllEfforts(string slug = "test-model")
        {
            return Catalog(Entry(slug, 1,
                new[] { "minimal", "low", "medium", "high", "xhigh", "max" }));
        }

        /// <summary>
        /// Trimmed copy of a live <c>/backend-api/codex/models</c> response
        /// (client_version=0.155.1, 2026-09-22), keeping a few of the Codex-only
        /// fields the parser must ignore.
        /// </summary>
        public const string LiveModelsResponse = @"{
  ""models"": [
    {
      ""slug"": ""gpt-6-astra"",
      ""display_name"": ""GPT-6-Astra"",
      ""description"": ""Our most capable model for complex, demanding work."",
      ""default_reasoning_level"": ""medium"",
      ""supported_reasoning_levels"": [
        { ""effort"": ""low"", ""description"": ""Fast responses with lighter reasoning"" },
        { ""effort"": ""medium"", ""description"": ""Balances speed and reasoning depth for everyday tasks"" },
        { ""effort"": ""high"", ""description"": ""Greater reasoning depth for complex problems"" },
        { ""effort"": ""xhigh"", ""description"": ""Extra high reasoning depth for complex problems"" },
        { ""effort"": ""max"", ""description"": ""Maximum reasoning depth for the hardest problems"" },
        { ""effort"": ""ultra"", ""description"": ""Maximum reasoning with automatic task delegation"" }
      ],
      ""shell_type"": ""unified_exec"",
      ""visibility"": ""list"",
      ""supported_in_api"": true,
      ""priority"": 1,
      ""upgrade"": null,
      ""model_messages"": ""(long Codex prompt text)"",
      ""tool_mode"": ""code_mode_only"",
      ""context_window"": 272000
    },
    {
      ""slug"": ""gpt-reserve"",
      ""display_name"": ""GPT-Reserve"",
      ""description"": ""Fast and affordable agentic coding model."",
      ""supported_reasoning_levels"": [
        { ""effort"": ""low"", ""description"": ""x"" },
        { ""effort"": ""medium"", ""description"": ""x"" }
      ],
      ""visibility"": ""hide"",
      ""priority"": 3,
      ""upgrade"": null
    },
    {
      ""slug"": ""gpt-5.6-sol"",
      ""display_name"": ""GPT-5.6-Sol"",
      ""description"": ""Reliable agentic workhorse for everyday tasks."",
      ""supported_reasoning_levels"": [
        { ""effort"": ""low"", ""description"": ""x"" },
        { ""effort"": ""medium"", ""description"": ""x"" },
        { ""effort"": ""high"", ""description"": ""x"" },
        { ""effort"": ""xhigh"", ""description"": ""x"" },
        { ""effort"": ""max"", ""description"": ""x"" },
        { ""effort"": ""ultra"", ""description"": ""x"" }
      ],
      ""visibility"": ""list"",
      ""priority"": 4,
      ""upgrade"": {
        ""model"": ""gpt-6-sol"",
        ""migration_markdown"": ""Meet GPT-6 Sol"",
        ""retirement_at"": null
      }
    },
    {
      ""slug"": ""gpt-5.5"",
      ""display_name"": ""GPT-5.5"",
      ""description"": ""Proven previous-generation model for coding and general work."",
      ""supported_reasoning_levels"": [
        { ""effort"": ""low"", ""description"": ""x"" },
        { ""effort"": ""medium"", ""description"": ""x"" },
        { ""effort"": ""high"", ""description"": ""x"" },
        { ""effort"": ""xhigh"", ""description"": ""x"" }
      ],
      ""visibility"": ""list"",
      ""priority"": 12,
      ""upgrade"": {
        ""model"": ""gpt-5.6-sol"",
        ""migration_markdown"": ""GPT-5.5 retires on October 14, 2026."",
        ""retirement_at"": ""2026-10-14T19:00:00Z""
      }
    }
  ]
}";

        /// <summary>Verbatim 400 body from /responses for an invalid effort (2026-09-22).</summary>
        public const string InvalidEffortErrorBody =
            "{\n  \"error\": {\n    \"message\": \"Invalid value: 'outlookai_probe'. Supported values are: "
            + "'none', 'minimal', 'low', 'medium', 'high', 'xhigh', and 'max'.\",\n"
            + "    \"type\": \"invalid_request_error\",\n    \"param\": \"reasoning.effort\",\n"
            + "    \"code\": \"invalid_value\"\n  }\n}";
    }
}
