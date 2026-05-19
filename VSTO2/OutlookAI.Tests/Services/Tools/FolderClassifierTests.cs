using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class FolderClassifierTests
    {
        private readonly IFolderClassifier _classifier = new FolderClassifier();

        [Theory]
        [InlineData("Deleted Items")]
        [InlineData("Junk E-mail")]      // real Outlook name, with hyphen
        [InlineData("Junk Email")]       // older locale variant
        [InlineData("Drafts")]
        [InlineData("Outbox")]
        [InlineData("Sync Issues")]
        [InlineData("Sync Issues (This computer only)")]
        [InlineData("Conflicts")]
        [InlineData("Local Failures")]
        [InlineData("Server Failures")]
        [InlineData("RSS Feeds")]
        [InlineData("RSS Subscriptions")]
        [InlineData("Conversation Action Settings")]
        [InlineData("Conversation History")]
        [InlineData("Quick Step Settings")]
        [InlineData("News Feed")]
        [InlineData("Feeds")]
        [InlineData("Files")]
        [InlineData("Detected Items")]
        [InlineData("Working Set")]
        [InlineData("Yammer Root")]
        public void IsSystemFolder_KnownSystemNames_ReturnsTrue(string name)
        {
            Assert.True(_classifier.IsSystemFolder(name, defaultItemTypeIsMail: true));
        }

        [Theory]
        [InlineData("Inbox")]
        [InlineData("Archive")]
        [InlineData("Sent Items")]
        [InlineData("Projects")]
        [InlineData("Receipts")]
        [InlineData("Important Office emails")]
        public void IsSystemFolder_UserFolders_ReturnsFalse(string name)
        {
            Assert.False(_classifier.IsSystemFolder(name, defaultItemTypeIsMail: true));
        }

        [Fact]
        public void IsSystemFolder_NonMailDefaultItemType_AlwaysTrue()
        {
            Assert.True(_classifier.IsSystemFolder("Calendar", defaultItemTypeIsMail: false));
            Assert.True(_classifier.IsSystemFolder("Inbox", defaultItemTypeIsMail: false));
        }

        [Fact]
        public void IsSystemFolder_NullOrEmpty_ReturnsTrue()
        {
            Assert.True(_classifier.IsSystemFolder(null, defaultItemTypeIsMail: true));
            Assert.True(_classifier.IsSystemFolder("", defaultItemTypeIsMail: true));
        }

        [Fact]
        public void IsSystemFolder_NameMatchIsCaseInsensitive()
        {
            Assert.True(_classifier.IsSystemFolder("JUNK E-MAIL", defaultItemTypeIsMail: true));
            Assert.True(_classifier.IsSystemFolder("conversation history", defaultItemTypeIsMail: true));
        }
    }
}
