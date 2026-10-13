using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.QuickTools;
using Xunit;

namespace OutlookAI.Tests.Services.QuickTools
{
    public class PagePlanTests
    {
        private const string Keystone =
            "[[0.294,0.117],[0.719,0.158],[0.806,0.9],[0.206,0.842]]";

        private static string Json(string corners, string kind = "document_page", string color = "grayscale", string strength = "normal") =>
            "{\"kind\":\"" + kind + "\",\"reason\":\" lease page 1 \",\"corners\":" + corners
            + ",\"color\":\"" + color + "\",\"strength\":\"" + strength + "\"}";

        [Fact]
        public void FromJson_ValidPlan_IsAnalyzed()
        {
            var plan = PagePlan.FromJson(Json(Keystone, color: "color", strength: "strong"));

            Assert.True(plan.Analyzed);
            Assert.Equal(PageKind.DocumentPage, plan.Kind);
            Assert.Equal("lease page 1", plan.Reason);
            Assert.Equal(PageColor.Color, plan.Color);
            Assert.Equal(CleanupStrength.Strong, plan.Strength);
            Assert.Equal(0.294, plan.Corners[0].X, 9);
            Assert.Equal(0.842, plan.Corners[3].Y, 9);
        }

        [Fact]
        public void FromJson_SlightlyOutside_IsClamped()
        {
            var plan = PagePlan.FromJson(Json("[[-0.01,0],[1.01,0],[1,1],[0,1]]"));

            Assert.True(plan.Analyzed);
            Assert.Equal(0.0, plan.Corners[0].X, 9);
            Assert.Equal(1.0, plan.Corners[1].X, 9);
        }

        [Fact]
        public void FromJson_OutOfRange_UsesTheFullImage_AndKeepsTheRest()
        {
            var plan = PagePlan.FromJson(Json("[[0,0],[1.2,0],[1,1],[0,1]]", color: "color"));

            Assert.False(plan.Analyzed);
            Assert.Equal(PageColor.Color, plan.Color);
            Assert.Equal(PagePlan.FullImage.Select(p => p.X), plan.Corners.Select(p => p.X));
        }

        [Fact]
        public void FromJson_MirroredCorners_UseTheFullImage()
        {
            var plan = PagePlan.FromJson(Json("[[0,0],[0,1],[1,1],[1,0]]"));

            Assert.False(plan.Analyzed);
            Assert.Equal(1.0, plan.Corners[1].X, 9);
            Assert.Equal(0.0, plan.Corners[1].Y, 9);
        }

        [Fact]
        public void FromJson_TinyOutline_UsesTheFullImage()
        {
            var plan = PagePlan.FromJson(Json("[[0.4,0.4],[0.5,0.4],[0.5,0.5],[0.4,0.5]]"));

            Assert.False(plan.Analyzed);
        }

        [Fact]
        public void FromJson_NotADocument_KeepsItsReason()
        {
            var plan = PagePlan.FromJson(Json("[[0,0],[1,0],[1,1],[0,1]]", kind: "not_document"));

            Assert.True(plan.Analyzed);
            Assert.Equal(PageKind.NotDocument, plan.Kind);
            Assert.Equal("lease page 1", plan.Reason);
        }

        [Theory]
        [InlineData("not json")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("{\"kind\":\"maybe\",\"reason\":\"\",\"corners\":[[0,0],[1,0],[1,1],[0,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}")]
        [InlineData("{\"kind\":\"document_page\",\"reason\":\"\",\"corners\":[[0,0],[1,0],[1,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}")]
        public void FromJson_Unusable_IsTheFallback(string text)
        {
            var plan = PagePlan.FromJson(text);

            Assert.False(plan.Analyzed);
            Assert.Equal(PageKind.DocumentPage, plan.Kind);
            Assert.Equal(PageColor.Grayscale, plan.Color);
            Assert.Equal(CleanupStrength.Normal, plan.Strength);
        }

        [Fact]
        public void FromJson_ToleratesAFenceAroundTheJson()
        {
            var plan = PagePlan.FromJson("Here you go:\n```json\n" + Json(Keystone) + "\n```");

            Assert.True(plan.Analyzed);
        }

        // Strict structured outputs need every property listed and nothing extra.
        [Fact]
        public void Schema_IsStrictFriendly()
        {
            var schema = PagePlan.Schema();

            Assert.False((bool)schema["additionalProperties"]);
            var required = ((JArray)schema["required"]).Select(t => (string)t).OrderBy(s => s);
            var properties = ((JObject)schema["properties"]).Properties().Select(p => p.Name).OrderBy(s => s);
            Assert.Equal(properties, required);
        }
    }
}
