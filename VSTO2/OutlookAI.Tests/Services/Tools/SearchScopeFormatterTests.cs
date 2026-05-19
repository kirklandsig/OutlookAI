using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchScopeFormatterTests
    {
        [Fact]
        public void Format_SinglePath_QuotesIt()
        {
            var s = SearchScopeFormatter.Format(new[] { @"\\Mailbox - User\Inbox" });
            Assert.Equal(@"'\\Mailbox - User\Inbox'", s);
        }

        [Fact]
        public void Format_MultiplePaths_CommaSeparatedQuoted()
        {
            var s = SearchScopeFormatter.Format(new[]
            {
                @"\\Mailbox - User\Inbox",
                @"\\Archive PST\Sent Items",
            });
            Assert.Equal(@"'\\Mailbox - User\Inbox','\\Archive PST\Sent Items'", s);
        }

        [Fact]
        public void Format_PathContainingSingleQuote_IsDoubledForOutlook()
        {
            var s = SearchScopeFormatter.Format(new[] { @"\\Mailbox\O'Brien Folder" });
            // Outlook DASL scope escapes ' as ''
            Assert.Equal(@"'\\Mailbox\O''Brien Folder'", s);
        }

        [Fact]
        public void Format_NullOrEmpty_ReturnsEmptyString()
        {
            Assert.Equal("", SearchScopeFormatter.Format(null));
            Assert.Equal("", SearchScopeFormatter.Format(new string[0]));
        }

        [Fact]
        public void Format_SkipsNullAndWhitespacePaths()
        {
            var s = SearchScopeFormatter.Format(new[]
            {
                @"\\A\Inbox",
                null,
                "   ",
                @"\\B\Inbox",
            });
            Assert.Equal(@"'\\A\Inbox','\\B\Inbox'", s);
        }
    }
}
