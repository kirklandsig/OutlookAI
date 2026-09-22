using System;
using System.Linq;
using OutlookAI.Services.Models;
using Xunit;
using static OutlookAI.Tests.Helpers.TestCatalogs;

namespace OutlookAI.Tests.Services.Models
{
    public class ModelCatalogTests
    {
        private static readonly string[] Standard = { "low", "medium", "high", "xhigh" };

        [Fact]
        public void Constructor_SortsByPriorityThenSlug_AndDropsDuplicateSlugs()
        {
            var catalog = Catalog(
                Entry("zeta", 5, Standard),
                Entry("alpha", 5, Standard),
                Entry("first", 1, Standard),
                Entry("FIRST", 9, new[] { "low" }));

            Assert.Equal(new[] { "first", "alpha", "zeta" }, catalog.Models.Select(m => m.Slug));
            // The first occurrence wins; the later case-variant duplicate is dropped.
            Assert.Equal(Standard, catalog.Find("first").Efforts);
        }

        [Fact]
        public void Constructor_RequiresAtLeastOneModel()
        {
            Assert.Throws<ArgumentException>(() => Catalog());
        }

        [Fact]
        public void ListedSlugs_ExcludesHiddenModels_ButFindStillResolvesThem()
        {
            var catalog = Catalog(
                Entry("shown", 1, Standard),
                Entry("hidden", 2, Standard, listed: false));

            Assert.Equal(new[] { "shown" }, catalog.ListedSlugs);
            Assert.NotNull(catalog.Find("hidden"));
        }

        [Fact]
        public void Find_IsCaseInsensitive_AndReturnsCanonicalEntry()
        {
            var catalog = Catalog(Entry("gpt-5.6-sol", 1, Standard));

            Assert.Equal("gpt-5.6-sol", catalog.Find("GPT-5.6-SOL").Slug);
            Assert.Equal("gpt-5.6-sol", catalog.Find("  gpt-5.6-sol ").Slug);
            Assert.Null(catalog.Find(null));
            Assert.Null(catalog.Find("gpt-9"));
        }

        [Fact]
        public void EffortsFor_PutsNoneFirst_AndDisplayCasesKnownEfforts()
        {
            var catalog = Catalog(Entry("m", 1, new[] { "low", "medium", "high", "xhigh", "max" }));

            Assert.Equal(new[] { "None", "Low", "Medium", "High", "XHigh", "Max" }, catalog.EffortsFor("m"));
        }

        [Fact]
        public void EffortsFor_DropsCatalogEffortsTheServerRejects()
        {
            // 'ultra' is a Codex-client mode; the probed server set lacks it.
            var catalog = new ModelCatalog(
                new[] { Entry("m", 1, new[] { "low", "max", "ultra" }) },
                new[] { "none", "minimal", "low", "medium", "high", "xhigh", "max" },
                ModelCatalog.SourceChatGpt, FixedNow, "0.155.1");

            Assert.Equal(new[] { "None", "Low", "Max" }, catalog.EffortsFor("m"));
        }

        [Fact]
        public void EffortsFor_WithoutProbedServerSet_FallsBackToKnownServerEfforts()
        {
            var catalog = Catalog(Entry("m", 1, new[] { "low", "max", "ultra" }));

            Assert.Null(catalog.ServerEfforts);
            Assert.Equal(new[] { "None", "Low", "Max" }, catalog.EffortsFor("m"));
        }

        [Fact]
        public void EffortsFor_OffersNewServerEffort_WithoutAnAppUpdate()
        {
            var catalog = new ModelCatalog(
                new[] { Entry("m", 1, new[] { "high", "extreme" }) },
                new[] { "high", "extreme" },
                ModelCatalog.SourceChatGpt, FixedNow, "0.160.0");

            Assert.Equal(new[] { "None", "High", "Extreme" }, catalog.EffortsFor("m"));
        }

        [Fact]
        public void EffortsFor_FoldsCatalogNoneIntoTheAppNoneOption()
        {
            var catalog = Catalog(Entry("m", 1, new[] { "none", "low" }));

            Assert.Equal(new[] { "None", "Low" }, catalog.EffortsFor("m"));
        }

        [Fact]
        public void EffortsFor_UnknownModel_OffersOnlyNone()
        {
            var catalog = Catalog(Entry("m", 1, Standard));

            Assert.Equal(new[] { "None" }, catalog.EffortsFor("gpt-4.1-nano"));
            Assert.Equal(new[] { "None" }, catalog.EffortsFor(null));
        }

        [Fact]
        public void EffortsFor_ReturnsACopy()
        {
            var catalog = Catalog(Entry("m", 1, Standard));

            catalog.EffortsFor("m")[0] = "tampered";

            Assert.Equal("None", catalog.EffortsFor("m")[0]);
        }

        [Fact]
        public void AllOfferedEfforts_IsNoneThenUnionInFirstSeenOrder()
        {
            var catalog = Catalog(
                Entry("a", 1, new[] { "low", "medium" }),
                Entry("b", 2, new[] { "medium", "max", "ultra" }));

            Assert.Equal(new[] { "None", "Low", "Medium", "Max" }, catalog.AllOfferedEfforts);
        }

        [Theory]
        [InlineData("max", "Max")]
        [InlineData("XHIGH", "XHigh")]
        [InlineData(" Low ", "Low")]
        [InlineData("none", "None")]
        [InlineData("Ultra", null)]
        [InlineData("Extreme", null)]
        [InlineData("Minimal", null)]
        [InlineData("", null)]
        [InlineData(null, null)]
        public void NormalizeEffort_AcceptsNoneAndEffortsSomeModelOffers(string raw, string expected)
        {
            var catalog = Catalog(
                Entry("a", 1, Standard),
                Entry("b", 2, new[] { "low", "max", "ultra" }));

            Assert.Equal(expected, catalog.NormalizeEffort(raw));
        }

        [Fact]
        public void DefaultModelAt_IsFirstListedModelThatIsNotRetired()
        {
            var retired = new ModelUpgradeInfo("next", FixedNow.AddDays(-1));
            var catalog = Catalog(
                Entry("hidden-top", 0, Standard, listed: false),
                Entry("old", 1, Standard, upgrade: retired),
                Entry("next", 2, Standard));

            Assert.Equal("next", catalog.DefaultModelAt(FixedNow));
            Assert.Equal("old", catalog.DefaultModelAt(FixedNow.AddDays(-2)));
        }

        [Fact]
        public void DefaultModelAt_EverythingRetired_FallsBackToFirstListed()
        {
            var catalog = Catalog(
                Entry("a", 1, Standard, upgrade: new ModelUpgradeInfo("b", FixedNow.AddDays(-1))),
                Entry("b", 2, Standard, upgrade: new ModelUpgradeInfo("a", FixedNow.AddDays(-1))));

            Assert.Equal("a", catalog.DefaultModelAt(FixedNow));
        }

        [Fact]
        public void ResolveEffectiveModel_KeepsConfiguredModel_WhenOfferedAndNotRetired()
        {
            var catalog = Catalog(Entry("top", 1, Standard), Entry("mine", 2, Standard));

            Assert.Equal("mine", catalog.ResolveEffectiveModel("MINE", FixedNow));
        }

        [Fact]
        public void ResolveEffectiveModel_SwitchesToUpgradeTarget_OnlyOnceRetirementPasses()
        {
            var retireAt = new DateTimeOffset(2026, 10, 14, 19, 0, 0, TimeSpan.Zero);
            var catalog = Catalog(
                Entry("top", 1, Standard),
                Entry("successor", 2, Standard),
                Entry("old", 3, Standard, upgrade: new ModelUpgradeInfo("successor", retireAt)));

            Assert.Equal("old", catalog.ResolveEffectiveModel("old", retireAt.AddSeconds(-1)));
            Assert.Equal("successor", catalog.ResolveEffectiveModel("old", retireAt));
        }

        [Fact]
        public void ResolveEffectiveModel_SoftUpgradeWithoutDate_KeepsConfiguredModel()
        {
            var catalog = Catalog(
                Entry("newer", 1, Standard),
                Entry("current", 2, Standard, upgrade: new ModelUpgradeInfo("newer", null)));

            Assert.Equal("current", catalog.ResolveEffectiveModel("current", FixedNow));
        }

        [Fact]
        public void ResolveEffectiveModel_FollowsARetirementChain()
        {
            var past = FixedNow.AddDays(-1);
            var catalog = Catalog(
                Entry("top", 1, Standard),
                Entry("c", 2, Standard),
                Entry("b", 3, Standard, upgrade: new ModelUpgradeInfo("c", past)),
                Entry("a", 4, Standard, upgrade: new ModelUpgradeInfo("b", past)));

            Assert.Equal("c", catalog.ResolveEffectiveModel("a", FixedNow));
        }

        [Fact]
        public void ResolveEffectiveModel_CycleOrMissingTarget_FallsBackToDefault()
        {
            var past = FixedNow.AddDays(-1);
            var catalog = Catalog(
                Entry("top", 1, Standard),
                Entry("x", 2, Standard, upgrade: new ModelUpgradeInfo("y", past)),
                Entry("y", 3, Standard, upgrade: new ModelUpgradeInfo("x", past)),
                Entry("orphan", 4, Standard, upgrade: new ModelUpgradeInfo("gone", past)),
                Entry("dead-end", 5, Standard, upgrade: new ModelUpgradeInfo(null, past)));

            Assert.Equal("top", catalog.ResolveEffectiveModel("x", FixedNow));
            Assert.Equal("top", catalog.ResolveEffectiveModel("orphan", FixedNow));
            Assert.Equal("top", catalog.ResolveEffectiveModel("dead-end", FixedNow));
        }

        [Fact]
        public void ResolveEffectiveModel_ModelNoLongerOffered_UsesDefault()
        {
            var catalog = Catalog(Entry("top", 1, Standard));

            Assert.Equal("top", catalog.ResolveEffectiveModel("gpt-5.4", FixedNow));
            Assert.Equal("top", catalog.ResolveEffectiveModel("", FixedNow));
            Assert.Equal("top", catalog.ResolveEffectiveModel(null, FixedNow));
        }

        [Theory]
        [InlineData("None", null)]
        [InlineData("none", null)]
        [InlineData("", null)]
        [InlineData(null, null)]
        [InlineData("XHigh", "xhigh")]
        [InlineData("max", "max")]
        [InlineData("Minimal", null)]   // not offered for this model -> omit
        [InlineData("Ultra", null)]     // client-only mode -> omit
        public void ResolveWireEffort_MapsOfferedEffortsAndOmitsTheRest(string effort, string expectedWire)
        {
            var catalog = Catalog(Entry("m", 1, new[] { "low", "xhigh", "max", "ultra" }));

            Assert.Equal(expectedWire, catalog.ResolveWireEffort("m", effort));
        }

        [Fact]
        public void ResolveWireEffort_UnknownModel_PassesEffortThroughLowercased()
        {
            var catalog = Catalog(Entry("m", 1, Standard));

            Assert.Equal("high", catalog.ResolveWireEffort("custom-model", "High"));
        }

        [Fact]
        public void Newest_PicksTheLatestData_AndPrefersTheFirstOnATie()
        {
            ModelCatalog At(string slug, DateTimeOffset? fetchedAt) => new ModelCatalog(
                new[] { Entry(slug, 1, Standard) }, null, ModelCatalog.SourceChatGpt, fetchedAt, "0.155.1");

            var older = At("older", FixedNow.AddDays(-3));
            var newer = At("newer", FixedNow);
            var tie = At("tie", FixedNow);

            Assert.Same(newer, ModelCatalog.Newest(older, null, newer));
            Assert.Same(newer, ModelCatalog.Newest(newer, tie));
            Assert.Same(older, ModelCatalog.Newest(At("undated", null), older));
            Assert.Null(ModelCatalog.Newest(null, null));
        }

        [Fact]
        public void Constructor_KeepsFetchMetadata()
        {
            var catalog = new ModelCatalog(new[] { Entry("m", 1, Standard) }, null,
                ModelCatalog.SourceChatGpt, FixedNow, "0.156.0", clientVersionSource: "config", accountId: "acct-9");

            Assert.Equal("config", catalog.ClientVersionSource);
            Assert.Equal("acct-9", catalog.AccountId);
        }

        [Fact]
        public void ModelUpgradeInfo_IsRetiredAt_RequiresADate()
        {
            var when = FixedNow;

            Assert.False(new ModelUpgradeInfo("x", null).IsRetiredAt(when));
            Assert.False(new ModelUpgradeInfo("x", when.AddTicks(1)).IsRetiredAt(when));
            Assert.True(new ModelUpgradeInfo("x", when).IsRetiredAt(when));
        }
    }
}
