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

        [Fact]
        public void MaxSearchFolders_IsNotCappedByListFoldersLimit()
        {
            Assert.Equal(200, SearchFallbackBudget.MaxListFolders);
            Assert.True(SearchFallbackBudget.MaxSearchFolders > SearchFallbackBudget.MaxListFolders);
            Assert.True(SearchFallbackBudget.MaxSearchFolders >= 5000);
        }

        [Fact]
        public void MaxFoldersForSearch_BroadNewestAllMailFinite_ReturnsInteractiveCap()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                SortOrder = "newest",
                MaxResults = 100,
            };

            Assert.Equal(200, SearchFallbackBudget.MaxFoldersForSearch(args, allMail: true));
        }

        [Theory]
        [InlineData("query")]
        [InlineData("from")]
        [InlineData("to")]
        [InlineData("subject")]
        [InlineData("body")]
        public void MaxFoldersForSearch_TargetedAllMailSearches_ReturnFullCap(string filter)
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                SortOrder = "newest",
                MaxResults = 100,
            };
            if (filter == "query") args.Query = "ein";
            if (filter == "from") args.From = "vendor@example.com";
            if (filter == "to") args.To = "Susan";
            if (filter == "subject") args.SubjectContains = "invoice";
            if (filter == "body") args.BodyContains = "tax id";

            Assert.Equal(SearchFallbackBudget.MaxSearchFolders, SearchFallbackBudget.MaxFoldersForSearch(args, allMail: true));
        }

        [Fact]
        public void MaxFoldersForSearch_CountMode_ReturnsFullCap()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                SortOrder = "newest",
                MaxResults = int.MaxValue,
            };

            Assert.Equal(SearchFallbackBudget.MaxSearchFolders, SearchFallbackBudget.MaxFoldersForSearch(args, allMail: true));
        }

        [Fact]
        public void MaxFoldersForSearch_Oldest_ReturnsFullCap()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                SortOrder = "oldest",
                MaxResults = 100,
            };

            Assert.Equal(SearchFallbackBudget.MaxSearchFolders, SearchFallbackBudget.MaxFoldersForSearch(args, allMail: true));
        }

        [Fact]
        public void MaxFoldersForSearch_ExplicitOrCurrentFolder_ReturnsFullCap()
        {
            Assert.Equal(SearchFallbackBudget.MaxSearchFolders, SearchFallbackBudget.MaxFoldersForSearch(
                new SearchMessagesArgs { Scope = "all_mail", FolderId = "folder-1", SortOrder = "newest", MaxResults = 100 },
                allMail: true));
            Assert.Equal(SearchFallbackBudget.MaxSearchFolders, SearchFallbackBudget.MaxFoldersForSearch(
                new SearchMessagesArgs { Scope = "current_folder", SortOrder = "newest", MaxResults = 100 },
                allMail: false));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_RecipientAllMailNewestAtMax_ReturnsTrue()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                To = "Susan",
                SortOrder = "newest",
                MaxResults = 25,
            };

            Assert.True(SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, "all_mail", candidateCount: 25));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_AutoWithoutFolderAtMax_ReturnsTrue()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "auto",
                To = "Susan",
                SortOrder = "newest",
                MaxResults = 25,
            };

            Assert.True(SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, "auto", candidateCount: 25));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_BeforeMax_ReturnsFalse()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                To = "Susan",
                SortOrder = "newest",
                MaxResults = 25,
            };

            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, "all_mail", candidateCount: 24));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_NonPositiveMaxResults_ReturnsFalse()
        {
            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(
                new SearchMessagesArgs { Scope = "all_mail", To = "Susan", SortOrder = "newest", MaxResults = 0 },
                "all_mail",
                candidateCount: 25));
            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(
                new SearchMessagesArgs { Scope = "all_mail", To = "Susan", SortOrder = "newest", MaxResults = -1 },
                "all_mail",
                candidateCount: 25));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_NonRecipientAllMail_ReturnsFalse()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                To = "",
                SortOrder = "newest",
                MaxResults = 25,
            };

            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, "all_mail", candidateCount: 25));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_ExplicitFolder_ReturnsFalse()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "auto",
                FolderId = "folder-1",
                To = "Susan",
                SortOrder = "newest",
                MaxResults = 25,
            };

            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, "auto", candidateCount: 25));
            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(
                new SearchMessagesArgs { Scope = "current_folder", To = "Susan", SortOrder = "newest", MaxResults = 25 },
                "current_folder",
                candidateCount: 25));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_OldestSort_ReturnsFalse()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                To = "Susan",
                SortOrder = "oldest",
                MaxResults = 25,
            };

            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, "all_mail", candidateCount: 25));
        }

        [Fact]
        public void ShouldStopRecipientAllMailScan_CountMode_ReturnsFalse()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                To = "Susan",
                SortOrder = "newest",
                MaxResults = int.MaxValue,
            };

            Assert.False(SearchFallbackBudget.ShouldStopRecipientAllMailScan(args, "all_mail", candidateCount: int.MaxValue));
        }

        [Fact]
        public void ShouldStopBroadAllMailScan_BroadNewestAllMailAtMax_ReturnsTrue()
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                SortOrder = "newest",
                MaxResults = 100,
            };

            Assert.True(SearchFallbackBudget.ShouldStopBroadAllMailScan(args, "all_mail", candidateCount: 100));
        }

        [Theory]
        [InlineData("query")]
        [InlineData("from")]
        [InlineData("to")]
        [InlineData("subject")]
        [InlineData("body")]
        public void ShouldStopBroadAllMailScan_TargetedAllMailAtMax_ReturnsFalse(string filter)
        {
            var args = new SearchMessagesArgs
            {
                Scope = "all_mail",
                SortOrder = "newest",
                MaxResults = 100,
            };
            if (filter == "query") args.Query = "ein";
            if (filter == "from") args.From = "vendor@example.com";
            if (filter == "to") args.To = "Susan";
            if (filter == "subject") args.SubjectContains = "invoice";
            if (filter == "body") args.BodyContains = "tax id";

            Assert.False(SearchFallbackBudget.ShouldStopBroadAllMailScan(args, "all_mail", candidateCount: 100));
        }
    }
}
