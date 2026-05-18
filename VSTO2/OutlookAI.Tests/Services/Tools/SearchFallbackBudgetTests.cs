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
        public void PerFolderItems_MaxResults1_ReturnsAtLeastFive()
        {
            var args = new SearchMessagesArgs { MaxResults = 1 };
            var limit = SearchFallbackBudget.PerFolderItems(args);
            Assert.True(limit >= 5);
            // Tight upper bound: we don't want to scan thousands of items
            // per folder just because the user asked for one result.
            Assert.True(limit <= 20);
        }

        [Fact]
        public void PerFolderItems_MaxResults25_BoundedAround125()
        {
            var args = new SearchMessagesArgs { MaxResults = 25 };
            var limit = SearchFallbackBudget.PerFolderItems(args);
            Assert.True(limit >= 25);
            Assert.True(limit <= 250);
        }

        [Fact]
        public void PerFolderItems_NullArgs_ReturnsSafeDefault()
        {
            var limit = SearchFallbackBudget.PerFolderItems(null);
            Assert.True(limit >= 5);
            Assert.True(limit <= 50);
        }

        [Fact]
        public void PerFolderItems_ZeroOrNegativeMaxResults_ReturnsFloor()
        {
            Assert.Equal(5, SearchFallbackBudget.PerFolderItems(new SearchMessagesArgs { MaxResults = 0 }));
            Assert.Equal(5, SearchFallbackBudget.PerFolderItems(new SearchMessagesArgs { MaxResults = -10 }));
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
