using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SenderKeyNormalizerTests
    {
        [Theory]
        [InlineData("Jane Doe", "jane@example.com", "Jane Doe")]
        [InlineData("Jane Doe", null,               "Jane Doe")]
        [InlineData("Jane Doe", "",                 "Jane Doe")]
        public void NameWins_WhenPresent(string name, string email, string expected)
        {
            Assert.Equal(expected, SenderKeyNormalizer.Normalize(name, email));
        }

        [Theory]
        [InlineData("", "Jane@Example.com", "jane@example.com")]
        [InlineData(null, "Bob@Example.com", "bob@example.com")]
        [InlineData("   ", "Carol@example.com", "carol@example.com")]
        public void EmailLowercased_WhenNameMissing(string name, string email, string expected)
        {
            Assert.Equal(expected, SenderKeyNormalizer.Normalize(name, email));
        }

        [Theory]
        [InlineData("  Jane Doe  ", "jane@example.com", "Jane Doe")]
        public void NameTrimmed(string name, string email, string expected)
        {
            Assert.Equal(expected, SenderKeyNormalizer.Normalize(name, email));
        }

        [Fact]
        public void BothMissing_ReturnsUnknownSender()
        {
            Assert.Equal("(unknown sender)", SenderKeyNormalizer.Normalize(null, null));
            Assert.Equal("(unknown sender)", SenderKeyNormalizer.Normalize("", ""));
            Assert.Equal("(unknown sender)", SenderKeyNormalizer.Normalize("   ", "   "));
        }
    }
}
