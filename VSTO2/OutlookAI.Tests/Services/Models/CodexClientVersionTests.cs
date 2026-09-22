using OutlookAI.Services.Models;
using Xunit;

namespace OutlookAI.Tests.Services.Models
{
    public class CodexClientVersionTests
    {
        [Theory]
        [InlineData("rust-v0.155.1", "0.155.1")]   // openai/codex GitHub tag shape
        [InlineData("v1.2.3", "1.2.3")]
        [InlineData("0.155.1", "0.155.1")]
        [InlineData(" 0.160.0 ", "0.160.0")]
        [InlineData("0.155.1-alpha.2", "0.155.1")]
        [InlineData("0.155.1+build.7", "0.155.1")]
        [InlineData("0.155", null)]                // server rejects non-triples
        [InlineData("1.2.3.4", null)]
        [InlineData("abc", null)]
        [InlineData("rust-v", null)]
        [InlineData("99999999999.0.0", null)]      // overflow
        [InlineData("", null)]
        [InlineData(null, null)]
        public void Normalize_KeepsOnlyMajorMinorPatch(string raw, string expected)
        {
            Assert.Equal(expected, CodexClientVersion.Normalize(raw));
        }

        [Fact]
        public void Max_ComparesNumerically_AndSkipsInvalidCandidates()
        {
            Assert.Equal("0.156.0", CodexClientVersion.Max("0.155.1", null, "rust-v0.156.0", "junk"));
            Assert.Equal("0.100.0", CodexClientVersion.Max("0.99.0", "0.100.0"));
            Assert.Equal("1.0.0", CodexClientVersion.Max("0.999.999", "1.0.0"));
            Assert.Null(CodexClientVersion.Max(null, "junk"));
        }

        [Fact]
        public void BuiltInFloor_IsAValidVersion()
        {
            Assert.Equal(CodexClientVersion.BuiltInFloor, CodexClientVersion.Normalize(CodexClientVersion.BuiltInFloor));
        }
    }
}
