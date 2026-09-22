using System;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using OutlookAI.Services.Models;
using Xunit;
using static OutlookAI.Tests.Helpers.TestCatalogs;

namespace OutlookAI.Tests.Services.Models
{
    public class ModelCatalogJsonTests
    {
        [Fact]
        public void ParseModelsResponse_ReadsTheLiveShape_AndIgnoresCodexOnlyFields()
        {
            var models = ModelCatalogJson.ParseModelsResponse(LiveModelsResponse);

            Assert.Equal(new[] { "gpt-6-astra", "gpt-reserve", "gpt-5.6-sol", "gpt-5.5" }, models.Select(m => m.Slug));

            var astra = models[0];
            Assert.Equal("GPT-6-Astra", astra.DisplayName);
            Assert.Equal("Our most capable model for complex, demanding work.", astra.Description);
            Assert.True(astra.Listed);
            Assert.Equal(1, astra.Priority);
            Assert.Equal(new[] { "low", "medium", "high", "xhigh", "max", "ultra" }, astra.Efforts);
            Assert.Null(astra.Upgrade);

            Assert.False(models[1].Listed);

            var sol = models[2];
            Assert.Equal("gpt-6-sol", sol.Upgrade.Model);
            Assert.Null(sol.Upgrade.RetirementAt);

            var gpt55 = models[3];
            Assert.Equal("gpt-5.6-sol", gpt55.Upgrade.Model);
            Assert.Equal(new DateTimeOffset(2026, 10, 14, 19, 0, 0, TimeSpan.Zero), gpt55.Upgrade.RetirementAt);
        }

        [Fact]
        public void ParseModelsResponse_SkipsUnsafeOrMissingSlugs_AndDuplicates()
        {
            var json = "{\"models\":["
                + "{\"slug\":\"good-1\"},"
                + "{\"slug\":\"has space\"},"
                + "{\"slug\":\"../evil\"},"
                + "{\"slug\":\"\"},"
                + "{\"display_name\":\"no slug\"},"
                + "{\"slug\":\"" + new string('a', 101) + "\"},"
                + "{\"slug\":42},"
                + "\"not-an-object\","
                + "{\"slug\":\"GOOD-1\"},"
                + "{\"slug\":\"good.2_x:y\"}"
                + "]}";

            var models = ModelCatalogJson.ParseModelsResponse(json);

            Assert.Equal(new[] { "good-1", "good.2_x:y" }, models.Select(m => m.Slug));
        }

        [Fact]
        public void ParseModelsResponse_NormalizesEfforts_AndDropsInvalidOnes()
        {
            var json = "{\"models\":[{\"slug\":\"m\",\"supported_reasoning_levels\":["
                + "{\"effort\":\" High \"},"
                + "{\"effort\":\"high\"},"
                + "{\"effort\":\"BAD VALUE\"},"
                + "{\"effort\":\"<script>\"},"
                + "{\"description\":\"no effort\"},"
                + "\"max\","
                + "{\"effort\":\"x-high_2\"}"
                + "]}]}";

            var m = ModelCatalogJson.ParseModelsResponse(json).Single();

            Assert.Equal(new[] { "high", "max", "x-high_2" }, m.Efforts);
        }

        [Fact]
        public void ParseModelsResponse_DefaultsMissingFields()
        {
            var m = ModelCatalogJson.ParseModelsResponse("{\"models\":[{\"slug\":\"bare\"}]}").Single();

            Assert.Equal("bare", m.DisplayName);
            Assert.Equal("", m.Description);
            Assert.True(m.Listed);                  // only an explicit non-"list" visibility hides
            Assert.Equal(ModelCatalogJson.DefaultPriority, m.Priority);
            Assert.Empty(m.Efforts);
            Assert.Null(m.Upgrade);
        }

        [Theory]
        [InlineData("hide")]
        [InlineData("none")]
        [InlineData("experimental")]
        public void ParseModelsResponse_NonListVisibility_IsHidden(string visibility)
        {
            var m = ModelCatalogJson.ParseModelsResponse(
                "{\"models\":[{\"slug\":\"m\",\"visibility\":\"" + visibility + "\"}]}").Single();

            Assert.False(m.Listed);
        }

        [Fact]
        public void ParseModelsResponse_ClipsLongText()
        {
            var json = "{\"models\":[{\"slug\":\"m\",\"display_name\":\"" + new string('D', 500)
                + "\",\"description\":\"" + new string('d', 5000) + "\"}]}";

            var m = ModelCatalogJson.ParseModelsResponse(json).Single();

            Assert.Equal(ModelCatalogJson.MaxDisplayNameLength, m.DisplayName.Length);
            Assert.Equal(ModelCatalogJson.MaxDescriptionLength, m.Description.Length);
        }

        [Fact]
        public void ParseModelsResponse_CapsTheModelCount()
        {
            var sb = new StringBuilder("{\"models\":[");
            for (int i = 0; i < ModelCatalogJson.MaxModels + 50; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append("{\"slug\":\"m").Append(i).Append("\"}");
            }
            sb.Append("]}");

            Assert.Equal(ModelCatalogJson.MaxModels, ModelCatalogJson.ParseModelsResponse(sb.ToString()).Count);
        }

        [Fact]
        public void ParseModelsResponse_IgnoresUnsafeUpgradeTargets_AndBadDates()
        {
            var json = "{\"models\":["
                + "{\"slug\":\"a\",\"upgrade\":{\"model\":\"bad slug\",\"retirement_at\":\"2026-10-14T19:00:00Z\"}},"
                + "{\"slug\":\"b\",\"upgrade\":{\"model\":\"c\",\"retirement_at\":\"not a date\"}},"
                + "{\"slug\":\"d\",\"upgrade\":{}}"
                + "]}";

            var models = ModelCatalogJson.ParseModelsResponse(json);

            Assert.Null(models[0].Upgrade.Model);
            Assert.NotNull(models[0].Upgrade.RetirementAt);
            Assert.Equal("c", models[1].Upgrade.Model);
            Assert.Null(models[1].Upgrade.RetirementAt);
            Assert.Null(models[2].Upgrade);
        }

        [Theory]
        [InlineData("9999-12-31T00:00:00Z")]
        [InlineData("1999-12-31T00:00:00Z")]
        [InlineData("0001-01-01T00:00:00Z")]
        public void ParseModelsResponse_IgnoresImplausibleRetirementDates(string date)
        {
            // Out-of-range dates can't even be formatted under some calendars
            // (UmAlQura), which would break the Settings dialog.
            var m = ModelCatalogJson.ParseModelsResponse(
                "{\"models\":[{\"slug\":\"m\",\"upgrade\":{\"model\":\"n\",\"retirement_at\":\"" + date + "\"}}]}").Single();

            Assert.Equal("n", m.Upgrade.Model);
            Assert.Null(m.Upgrade.RetirementAt);
        }

        [Fact]
        public void ParseCache_RejectsAnImplausibleFetchDate()
        {
            Assert.Null(ModelCatalogJson.ParseCache(
                "{\"schema_version\":1,\"fetched_at\":\"9999-01-01T00:00:00Z\",\"models\":[{\"slug\":\"m\"}]}"));
        }

        [Fact]
        public void ParseModelsResponse_WrongShape_Throws()
        {
            Assert.Throws<FormatException>(() => ModelCatalogJson.ParseModelsResponse("{\"data\":[]}"));
            Assert.Throws<FormatException>(() => ModelCatalogJson.ParseModelsResponse("[]"));
            Assert.ThrowsAny<JsonException>(() => ModelCatalogJson.ParseModelsResponse("<html>"));
        }

        [Fact]
        public void Cache_RoundTripsEverythingTheAppUses()
        {
            var original = new ModelCatalog(
                ModelCatalogJson.ParseModelsResponse(LiveModelsResponse),
                new[] { "none", "low", "medium", "high", "xhigh", "max" },
                ModelCatalog.SourceChatGpt,
                new DateTimeOffset(2026, 9, 22, 19, 40, 5, TimeSpan.Zero),
                "0.155.1",
                clientVersionSource: "github",
                accountId: "acct-123");

            var copy = ModelCatalogJson.ParseCache(ModelCatalogJson.SerializeCache(original));

            Assert.NotNull(copy);
            Assert.Equal(original.Source, copy.Source);
            Assert.Equal(original.FetchedAt, copy.FetchedAt);
            Assert.Equal(original.ClientVersion, copy.ClientVersion);
            Assert.Equal("github", copy.ClientVersionSource);
            Assert.Equal("acct-123", copy.AccountId);
            Assert.Equal(original.ServerEfforts, copy.ServerEfforts);
            Assert.Equal(original.Models.Select(m => m.Slug), copy.Models.Select(m => m.Slug));
            for (int i = 0; i < original.Models.Count; i++)
            {
                var a = original.Models[i];
                var b = copy.Models[i];
                Assert.Equal(a.DisplayName, b.DisplayName);
                Assert.Equal(a.Description, b.Description);
                Assert.Equal(a.Listed, b.Listed);
                Assert.Equal(a.Priority, b.Priority);
                Assert.Equal(a.Efforts, b.Efforts);
                Assert.Equal(a.Upgrade?.Model, b.Upgrade?.Model);
                Assert.Equal(a.Upgrade?.RetirementAt, b.Upgrade?.RetirementAt);
            }
        }

        [Fact]
        public void Cache_RoundTripsAnUnprobedServerEffortSetAsNull()
        {
            var copy = ModelCatalogJson.ParseCache(ModelCatalogJson.SerializeCache(AllEfforts()));

            Assert.Null(copy.ServerEfforts);
        }

        [Theory]
        [InlineData("{\"schema_version\":2,\"fetched_at\":\"2026-09-22T00:00:00Z\",\"models\":[{\"slug\":\"m\"}]}")]
        [InlineData("{\"fetched_at\":\"2026-09-22T00:00:00Z\",\"models\":[{\"slug\":\"m\"}]}")]
        [InlineData("{\"schema_version\":1,\"fetched_at\":\"2026-09-22T00:00:00Z\",\"models\":[]}")]
        [InlineData("{\"schema_version\":1,\"fetched_at\":\"2026-09-22T00:00:00Z\",\"models\":[{\"slug\":\"m\",\"visibility\":\"hide\"}]}")]
        [InlineData("{\"schema_version\":1,\"models\":[{\"slug\":\"m\"}]}")]
        [InlineData("not json")]
        [InlineData("")]
        [InlineData(null)]
        public void ParseCache_RejectsUnusableFiles(string json)
        {
            Assert.Null(ModelCatalogJson.ParseCache(json));
        }

        [Fact]
        public void ParseCache_DropsInvalidServerEfforts_AndBadClientVersion()
        {
            var json = "{\"schema_version\":1,\"source\":\"chatgpt\",\"fetched_at\":\"2026-09-22T00:00:00Z\","
                + "\"client_version\":\"not-a-version\","
                + "\"server_reasoning_efforts\":[\"low\",\"BAD ONE\",\"max\",7],"
                + "\"models\":[{\"slug\":\"m\"}]}";

            var catalog = ModelCatalogJson.ParseCache(json);

            Assert.Equal(new[] { "low", "max" }, catalog.ServerEfforts);
            Assert.Null(catalog.ClientVersion);
        }

        [Fact]
        public void ParseSupportedEffortsFromError_ReadsTheServerEffortSet()
        {
            var efforts = ModelCatalogJson.ParseSupportedEffortsFromError(InvalidEffortErrorBody);

            Assert.Equal(new[] { "none", "minimal", "low", "medium", "high", "xhigh", "max" }, efforts);
        }

        [Theory]
        [InlineData("{\"detail\":\"The 'x' model is not supported when using Codex with a ChatGPT account.\"}")]
        [InlineData("{\"error\":{\"message\":\"Invalid value: 'x'. Supported values are: 'a', 'b'.\",\"param\":\"tools\"}}")]
        [InlineData("{\"error\":{\"message\":\"Something else went wrong.\",\"param\":\"reasoning.effort\"}}")]
        [InlineData("{\"error\":{\"message\":\"Invalid value: 'x'. Supported values are: 'low'.\",\"param\":\"reasoning.effort\"}}")]
        [InlineData("<html>502</html>")]
        [InlineData("")]
        [InlineData(null)]
        public void ParseSupportedEffortsFromError_ReturnsNull_ForAnythingElse(string body)
        {
            Assert.Null(ModelCatalogJson.ParseSupportedEffortsFromError(body));
        }
    }
}
