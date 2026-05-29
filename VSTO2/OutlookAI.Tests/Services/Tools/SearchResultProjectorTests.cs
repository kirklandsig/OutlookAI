using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchResultProjectorTests
    {
        private static MessageProjectionInput Item(
            string id, int year, string folder = "Inbox", Func<string> snippet = null)
        {
            return new MessageProjectionInput
            {
                Id = id,
                Subject = "s-" + id,
                From = "f-" + id,
                To = new[] { "to-" + id },
                ReceivedAt = new DateTimeOffset(year, 1, 1, 0, 0, 0, TimeSpan.Zero),
                HasAttachments = false,
                FolderName = folder,
                FolderDefaultItemTypeIsMail = true,
                SnippetFactory = snippet ?? (() => "snip-" + id),
            };
        }

        [Fact]
        public void Project_NewestSort_ReturnsByReceivedDesc()
        {
            var input = new[] { Item("a", 2010), Item("b", 2024), Item("c", 2017) };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(new[] { "b", "c", "a" }, result.Messages.Select(r => r.Id).ToArray());
        }

        [Fact]
        public void Project_OldestSort_ReturnsByReceivedAsc()
        {
            var input = new[] { Item("a", 2010), Item("b", 2024), Item("c", 2017) };
            var args = new SearchMessagesArgs { SortOrder = "oldest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(new[] { "a", "c", "b" }, result.Messages.Select(r => r.Id).ToArray());
        }

        [Fact]
        public void Project_TopN_ClampsToMaxResults()
        {
            var input = Enumerable.Range(0, 20).Select(i => Item("i" + i, 2000 + i)).ToList();
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 3 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(3, result.Messages.Count);
        }

        [Fact]
        public void Project_SkipsSystemFolders()
        {
            var input = new[]
            {
                Item("good",  2024, folder: "Inbox"),
                Item("junk",  2024, folder: "Junk E-mail"),
                Item("trash", 2024, folder: "Deleted Items"),
            };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(new[] { "good" }, result.Messages.Select(r => r.Id).ToArray());
        }

        [Fact]
        public void Project_DefersSnippet_OnlyForItemsInTopN()
        {
            int snippetCalls = 0;
            var input = Enumerable.Range(0, 10).Select(i =>
            {
                int captured = i;
                return new MessageProjectionInput
                {
                    Id = "i" + captured,
                    Subject = "s",
                    From = "f",
                    To = new string[0],
                    ReceivedAt = new DateTimeOffset(2000 + captured, 1, 1, 0, 0, 0, TimeSpan.Zero),
                    HasAttachments = false,
                    FolderName = "Inbox",
                    FolderDefaultItemTypeIsMail = true,
                    SnippetFactory = () => { snippetCalls++; return "snip" + captured; },
                };
            }).ToList();

            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 3 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());

            Assert.Equal(3, result.Messages.Count);
            Assert.Equal(3, snippetCalls); // exactly the surviving top-N had snippet evaluated
        }

        [Fact]
        public void Project_SnippetFactoryThrowing_DoesNotBreakBatch()
        {
            var input = new[]
            {
                new MessageProjectionInput {
                    Id = "ok",  Subject = "s", From = "f", To = new string[0],
                    ReceivedAt = new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero),
                    FolderName = "Inbox", FolderDefaultItemTypeIsMail = true,
                    SnippetFactory = () => "good",
                },
                new MessageProjectionInput {
                    Id = "bad", Subject = "s", From = "f", To = new string[0],
                    ReceivedAt = new DateTimeOffset(2023, 1, 1, 0, 0, 0, TimeSpan.Zero),
                    FolderName = "Inbox", FolderDefaultItemTypeIsMail = true,
                    SnippetFactory = () => { throw new InvalidOperationException("boom"); },
                },
            };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());
            Assert.Equal(2, result.Messages.Count);
            Assert.Equal("good", result.Messages.First(r => r.Id == "ok").Snippet);
            Assert.Equal("", result.Messages.First(r => r.Id == "bad").Snippet);
        }

        [Fact]
        public void Project_ReportsTotalMatchesAndTruncated_WhenClamped()
        {
            var input = Enumerable.Range(0, 10).Select(i => Item("i" + i, 2000 + i)).ToList();
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 3 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());

            Assert.Equal(3, result.Messages.Count);
            Assert.Equal(10, result.TotalMatches);
            Assert.True(result.Truncated);
        }

        [Fact]
        public void Project_NotTruncated_WhenUnderLimit()
        {
            var input = new[] { Item("a", 2010), Item("b", 2024) };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 25 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());

            Assert.Equal(2, result.Messages.Count);
            Assert.Equal(2, result.TotalMatches);
            Assert.False(result.Truncated);
        }

        [Fact]
        public void Project_TotalMatches_ExcludesSystemFolderItems()
        {
            // TotalMatches must count AFTER the system-folder filter, so junk /
            // deleted items never inflate the reported total (or the truncated flag).
            var input = new[]
            {
                Item("good",  2024, folder: "Inbox"),
                Item("junk",  2024, folder: "Junk E-mail"),
                Item("trash", 2024, folder: "Deleted Items"),
            };
            var args = new SearchMessagesArgs { SortOrder = "newest", MaxResults = 5 };
            var result = SearchResultProjector.Project(input, args, new FolderClassifier());

            Assert.Equal(1, result.Messages.Count);
            Assert.Equal(1, result.TotalMatches);
            Assert.False(result.Truncated);
        }
    }
}
