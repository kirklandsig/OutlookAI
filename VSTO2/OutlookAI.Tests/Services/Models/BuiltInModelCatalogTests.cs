using System;
using System.Linq;
using OutlookAI.Services.Models;
using Xunit;
using static OutlookAI.Tests.Helpers.TestCatalogs;

namespace OutlookAI.Tests.Services.Models
{
    public class BuiltInModelCatalogTests
    {
        private static ModelCatalog BuiltIn => BuiltInModelCatalog.Instance;

        [Fact]
        public void BuiltIn_IsMarkedBuiltIn_AndListsCurrentModels()
        {
            Assert.True(BuiltIn.IsBuiltIn);
            // Dated so a models.json older than this build loses to it.
            Assert.Equal(BuiltInModelCatalog.AsOf, BuiltIn.FetchedAt);
            Assert.Equal(CodexClientVersion.BuiltInFloor, BuiltIn.ClientVersion);
            Assert.Equal(
                new[] { "gpt-6-astra", "gpt-6-sol", "gpt-6-luna", "gpt-5.6-sol", "gpt-5.6-terra", "gpt-5.6-luna", "gpt-5.5" },
                BuiltIn.ListedSlugs);
        }

        [Fact]
        public void BuiltIn_DropsModelsChatGptAccountsCannotUse()
        {
            foreach (var stale in new[] { "gpt-5.5-pro", "gpt-5.4", "gpt-5.4-mini", "gpt-4.1-mini", "gpt-4.1-nano", "gpt-5.3-codex" })
            {
                Assert.Null(BuiltIn.Find(stale));
            }
        }

        [Fact]
        public void BuiltIn_DefaultIsTheTopPriorityModel()
        {
            Assert.Equal("gpt-6-astra", BuiltIn.DefaultModelAt(FixedNow));
        }

        [Fact]
        public void BuiltIn_Gpt55_RetiresToGpt56Sol_On2026_10_14()
        {
            var retirement = new DateTimeOffset(2026, 10, 14, 19, 0, 0, TimeSpan.Zero);

            Assert.Equal("gpt-5.5", BuiltIn.ResolveEffectiveModel("gpt-5.5", retirement.AddMinutes(-1)));
            Assert.Equal("gpt-5.6-sol", BuiltIn.ResolveEffectiveModel("gpt-5.5", retirement));
        }

        [Fact]
        public void BuiltIn_NeverOffersTheClientOnlyUltraMode()
        {
            foreach (var slug in BuiltIn.ListedSlugs)
            {
                Assert.DoesNotContain("Ultra", BuiltIn.EffortsFor(slug));
            }
        }

        [Fact]
        public void BuiltIn_EffortsMatchTheLiveProbe()
        {
            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh", "Max" }, BuiltIn.EffortsFor("gpt-6-astra"));
            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh" }, BuiltIn.EffortsFor("gpt-5.5"));
            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh", "Max" }, BuiltIn.AllOfferedEfforts);
        }

        [Fact]
        public void BuiltIn_SurvivesACacheRoundTrip()
        {
            // Built-in data goes through the same validation as a fetched catalog.
            var reparsed = ModelCatalogJson.ParseModelsResponse(
                "{\"models\":" + Newtonsoft.Json.Linq.JObject.Parse(ModelCatalogJson.SerializeCache(BuiltIn))["models"] + "}");

            Assert.Equal(BuiltIn.Models.Select(m => m.Slug), reparsed.Select(m => m.Slug));
            Assert.Equal(BuiltIn.Models.Select(m => string.Join(",", m.Efforts)), reparsed.Select(m => string.Join(",", m.Efforts)));
        }
    }
}
