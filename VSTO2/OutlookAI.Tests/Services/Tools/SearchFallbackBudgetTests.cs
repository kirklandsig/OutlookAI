using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    /// <summary>
    /// Phase 3b hotfix: bounds the per-folder iteration in the iterative
    /// fallback so a single huge folder cannot freeze Outlook for tens of
    /// seconds. The budget is pure logic separated for tests because the
    /// COM iteration that consumes it lives inside LiveOutlookSurface.
    /// </summary>
    public class SearchFallbackBudgetTests
    {
        [Fact]
        public void PerFolderItems_MaxResults1_ReturnsOne()
        {
            // For "oldest email" with MaxResults=1, taking ONE oldest item
            // per folder is exactly right. The global sort across folders
            // picks the absolute oldest. Anything > 1 is wasted property
            // access at ~100-300ms per item.
            var args = new SearchMessagesArgs { MaxResults = 1 };
            Assert.Equal(1, SearchFallbackBudget.PerFolderItems(args));
        }

        [Fact]
        public void PerFolderItems_MaxResults25_Returns25()
        {
            var args = new SearchMessagesArgs { MaxResults = 25 };
            Assert.Equal(25, SearchFallbackBudget.PerFolderItems(args));
        }

        [Fact]
        public void PerFolderItems_NullArgs_ReturnsOne()
        {
            Assert.Equal(1, SearchFallbackBudget.PerFolderItems(null));
        }

        [Fact]
        public void PerFolderItems_ZeroOrNegativeMaxResults_ReturnsOne()
        {
            Assert.Equal(1, SearchFallbackBudget.PerFolderItems(new SearchMessagesArgs { MaxResults = 0 }));
            Assert.Equal(1, SearchFallbackBudget.PerFolderItems(new SearchMessagesArgs { MaxResults = -10 }));
        }

        [Fact]
        public void PerFolderItems_CountModeIntMaxValue_UsesHardCap()
        {
            // CountMessages calls SearchMessages with MaxResults=int.MaxValue.
            // We still need a per-folder cap or Outlook freezes on huge
            // folders. 5000 is generous but bounded.
            var args = new SearchMessagesArgs { MaxResults = int.MaxValue };
            var limit = SearchFallbackBudget.PerFolderItems(args);
            Assert.Equal(5000, limit);
        }

        [Fact]
        public void DescendingForNewest_ReturnsTrue()
        {
            Assert.True(SearchFallbackBudget.DescendingForSortOrder("newest"));
            Assert.True(SearchFallbackBudget.DescendingForSortOrder(null));
            Assert.True(SearchFallbackBudget.DescendingForSortOrder(""));
            Assert.True(SearchFallbackBudget.DescendingForSortOrder("NEWEST"));
            Assert.True(SearchFallbackBudget.DescendingForSortOrder("bogus"));
        }

        [Fact]
        public void DescendingForOldest_ReturnsFalse()
        {
            Assert.False(SearchFallbackBudget.DescendingForSortOrder("oldest"));
            Assert.False(SearchFallbackBudget.DescendingForSortOrder("OLDEST"));
            Assert.False(SearchFallbackBudget.DescendingForSortOrder("Oldest"));
        }
    }
}
