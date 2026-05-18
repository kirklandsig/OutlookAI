using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchExecutionHelperTests
    {
        [Theory]
        [InlineData("oldest", false)]
        [InlineData("newest", true)]
        [InlineData(null, true)]
        [InlineData("bogus", true)]
        public void SortDescending_ReturnsExpectedDirection(string sortOrder, bool expected)
        {
            Assert.Equal(expected, LiveOutlookSurface.SortDescending(new SearchMessagesArgs { SortOrder = sortOrder }));
        }

        [Theory]
        [InlineData("Deleted Items", true)]
        [InlineData("Junk Email", true)]
        [InlineData("Drafts", true)]
        [InlineData("Outbox", true)]
        [InlineData("Sync Issues", true)]
        [InlineData("RSS Feeds", true)]
        [InlineData("Inbox", false)]
        [InlineData("Sent Items", false)]
        [InlineData("Archive", false)]
        [InlineData("Projects", false)]
        public void ShouldSkipAllMailFolder_ExcludesNoisyFolders(string name, bool expected)
        {
            Assert.Equal(expected, LiveOutlookSurface.ShouldSkipAllMailFolder(name));
        }

        [Fact]
        public void MergeAndSortSearchResults_HonorsOldestAndMaxResults()
        {
            var hits = new[]
            {
                new MessageSummary { Id = "new", ReceivedAt = DateTimeOffset.Parse("2024-01-01T00:00:00Z") },
                new MessageSummary { Id = "old", ReceivedAt = DateTimeOffset.Parse("2010-01-01T00:00:00Z") },
                new MessageSummary { Id = "mid", ReceivedAt = DateTimeOffset.Parse("2020-01-01T00:00:00Z") },
            };

            var merged = LiveOutlookSurface.MergeAndSortSearchResults(
                hits, new SearchMessagesArgs { SortOrder = "oldest", MaxResults = 2 });

            Assert.Equal(2, merged.Count);
            Assert.Equal("old", merged[0].Id);
            Assert.Equal("mid", merged[1].Id);
        }
    }
}
