using OutlookAI.Services.Variants;
using Xunit;

namespace OutlookAI.Tests.Services.Variants
{
    public class VariantParserTests
    {
        [Fact]
        public void Parse_WellFormedFencedJson_ReturnsVariants()
        {
            var input =
                "Here are three options:\n"
                + "```json\n"
                + "{\"variants\":["
                + "{\"tone\":\"Formal\",\"rationale\":\"r1\",\"subject\":\"s1\",\"body\":\"b1\"},"
                + "{\"tone\":\"Brief\",\"rationale\":\"r2\",\"subject\":\"s2\",\"body\":\"b2\"}"
                + "]}\n"
                + "```";
            var p = new VariantParser();

            var vs = p.Parse(input);

            Assert.Equal(2, vs.Count);
            Assert.Equal(Tone.Formal, vs[0].Tone);
            Assert.Equal("r1", vs[0].Rationale);
            Assert.Equal("s1", vs[0].Subject);
            Assert.Equal("b1", vs[0].Body);
            Assert.Equal(Tone.Brief, vs[1].Tone);
        }

        [Fact]
        public void Parse_FreeFormToneClampedToEnum()
        {
            var input = "```json\n{\"variants\":[{\"tone\":\"warmth\",\"body\":\"x\"}]}\n```";
            var p = new VariantParser();

            var vs = p.Parse(input);

            Assert.Single(vs);
            Assert.Equal(Tone.Friendly, vs[0].Tone);
            Assert.Equal("x", vs[0].Body);
            // Missing fields default to empty strings, not null.
            Assert.Equal("", vs[0].Rationale);
            Assert.Equal("", vs[0].Subject);
        }

        [Fact]
        public void Parse_BareJsonWithoutFence_IsAccepted()
        {
            var input = "{\"variants\":[{\"tone\":\"Direct\",\"body\":\"hello\"}]}";
            var p = new VariantParser();

            var vs = p.Parse(input);

            Assert.Single(vs);
            Assert.Equal(Tone.Direct, vs[0].Tone);
            Assert.Equal("hello", vs[0].Body);
        }

        [Fact]
        public void Parse_FenceWithoutLanguageTag_StillExtracts()
        {
            var input = "```\n{\"variants\":[{\"tone\":\"Formal\",\"body\":\"y\"}]}\n```";
            var p = new VariantParser();

            Assert.Single(p.Parse(input));
        }

        [Fact]
        public void Parse_MalformedJson_ReturnsEmptyList()
        {
            Assert.Empty(new VariantParser().Parse("not json"));
            Assert.Empty(new VariantParser().Parse(""));
            Assert.Empty(new VariantParser().Parse("   "));
            Assert.Empty(new VariantParser().Parse(null));
        }

        [Fact]
        public void Parse_MissingVariantsArray_ReturnsEmptyList()
        {
            var input = "```json\n{\"options\":[{\"tone\":\"Direct\"}]}\n```";
            Assert.Empty(new VariantParser().Parse(input));
        }

        [Fact]
        public void ClosestTo_KnownSynonyms_MapToExpectedEnumValues()
        {
            Assert.Equal(Tone.Friendly, ToneExtensions.ClosestTo("polite"));
            Assert.Equal(Tone.Friendly, ToneExtensions.ClosestTo("warmth"));
            Assert.Equal(Tone.Brief, ToneExtensions.ClosestTo("short and sweet"));
            Assert.Equal(Tone.Brief, ToneExtensions.ClosestTo("concise"));
            Assert.Equal(Tone.Persuasive, ToneExtensions.ClosestTo("sales pitch"));
            Assert.Equal(Tone.Apologetic, ToneExtensions.ClosestTo("apology"));
            Assert.Equal(Tone.Apologetic, ToneExtensions.ClosestTo("sorry"));
            Assert.Equal(Tone.Enthusiastic, ToneExtensions.ClosestTo("excited"));
            Assert.Equal(Tone.Technical, ToneExtensions.ClosestTo("engineer"));
            Assert.Equal(Tone.Formal, ToneExtensions.ClosestTo("OFFICIAL"));
            Assert.Equal(Tone.Diplomatic, ToneExtensions.ClosestTo("tactful"));
            // Unmapped string defaults to Direct.
            Assert.Equal(Tone.Direct, ToneExtensions.ClosestTo("zzz"));
            // Empty string also defaults to Direct.
            Assert.Equal(Tone.Direct, ToneExtensions.ClosestTo(""));
            Assert.Equal(Tone.Direct, ToneExtensions.ClosestTo(null));
        }
    }
}
