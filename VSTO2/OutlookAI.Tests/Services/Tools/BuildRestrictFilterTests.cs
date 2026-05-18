using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    /// <summary>
    /// Locks the DASL @SQL= filter syntax built from a SearchMessagesArgs.
    /// The Codex backend never sees this filter - it's the bridge between
    /// the model's structured request and Outlook's MAPI restrict API. If
    /// these clauses regress, search returns wrong results silently.
    /// </summary>
    public class BuildRestrictFilterTests
    {
        [Fact]
        public void EmptyArgs_ReturnsNullFilter()
        {
            var filter = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs());
            Assert.Null(filter);
        }

        [Fact]
        public void QueryOnly_BuildsSubjectOrBodyLike()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Query = "Q4" });
            Assert.StartsWith("@SQL=", f);
            Assert.Contains("urn:schemas:httpmail:subject LIKE '%Q4%'", f);
            Assert.Contains("urn:schemas:httpmail:textdescription LIKE '%Q4%'", f);
        }

        [Fact]
        public void From_MatchesDisplayNameOrEmail()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { From = "jane" });
            Assert.Contains("urn:schemas:httpmail:fromname LIKE '%jane%'", f);
            Assert.Contains("urn:schemas:httpmail:fromemail LIKE '%jane%'", f);
        }

        [Fact]
        public void SubjectContains_AddsSubjectLikeClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { SubjectContains = "plan" });
            Assert.Contains("urn:schemas:httpmail:subject LIKE '%plan%'", f);
            Assert.DoesNotContain("textdescription", f);
        }

        [Fact]
        public void BodyContains_AddsBodyLikeClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { BodyContains = "draft" });
            Assert.Contains("urn:schemas:httpmail:textdescription LIKE '%draft%'", f);
            Assert.DoesNotContain("subject LIKE", f);
        }

        [Fact]
        public void HasAttachment_TrueMapsToAttachmentClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { HasAttachment = true });
            Assert.Contains("urn:schemas:httpmail:hasattachment = 1", f);
        }

        [Fact]
        public void IsUnread_TrueMapsToUnreadClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { IsUnread = true });
            Assert.Contains("urn:schemas:httpmail:read = 0", f);
        }

        [Fact]
        public void IsFlagged_TrueUsesFlagStatusProperty()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { IsFlagged = true });
            Assert.Contains("0x10900003", f);
            Assert.Contains("= 2", f);
        }

        [Theory]
        [InlineData("low",    "= 0")]
        [InlineData("normal", "= 1")]
        [InlineData("high",   "= 2")]
        public void Importance_MapsToMapiPropertyValues(string ui, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Importance = ui });
            Assert.Contains("0x00170003", f);
            Assert.Contains(expected, f);
        }

        [Fact]
        public void Importance_UnknownValue_IsIgnored()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Importance = "Extreme" });
            Assert.Null(f);
        }

        [Fact]
        public void Compound_AllFieldsAndedTogether()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
            {
                Query = "Q4",
                From = "jane",
                IsUnread = true,
                HasAttachment = true,
                Importance = "high",
                DateFrom = new DateTimeOffset(2026, 5, 10, 0, 0, 0, TimeSpan.Zero),
            });
            Assert.StartsWith("@SQL=", f);
            // 6 clauses (query, from, unread, attachment, importance, datefrom)
            // joined by AND -> 5 separators.
            var ands = System.Text.RegularExpressions.Regex.Matches(f, " AND ").Count;
            Assert.Equal(5, ands);
        }

        [Fact]
        public void EscapesSingleQuotes()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Query = "Jane's Q4" });
            // Single quote inside DASL string is escaped by doubling.
            Assert.Contains("'%Jane''s Q4%'", f);
        }

        [Theory]
        [InlineData("any", null)]
        [InlineData("with", "urn:schemas:httpmail:hasattachment = 1")]
        [InlineData("without", "urn:schemas:httpmail:hasattachment = 0")]
        public void AttachmentFilter_MapsTriState(string value, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { AttachmentFilter = value });
            if (expected == null)
                Assert.Null(f);
            else
                Assert.Contains(expected, f);
        }

        [Theory]
        [InlineData("any", null)]
        [InlineData("unread", "urn:schemas:httpmail:read = 0")]
        [InlineData("read", "urn:schemas:httpmail:read = 1")]
        public void ReadStatus_MapsTriState(string value, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { ReadStatus = value });
            if (expected == null)
                Assert.Null(f);
            else
                Assert.Contains(expected, f);
        }

        [Theory]
        [InlineData("any", null)]
        [InlineData("flagged", "= 2")]
        [InlineData("unflagged", "<> 2")]
        public void FlagStatus_MapsTriState(string value, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { FlagStatus = value });
            if (expected == null)
                Assert.Null(f);
            else
            {
                Assert.Contains("0x10900003", f);
                Assert.Contains(expected, f);
            }
        }

        [Theory]
        [InlineData("any", null)]
        [InlineData("low", "= 0")]
        [InlineData("normal", "= 1")]
        [InlineData("high", "= 2")]
        public void ImportanceFilter_MapsTriState(string value, string expected)
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { ImportanceFilter = value });
            if (expected == null)
                Assert.Null(f);
            else
            {
                Assert.Contains("0x00170003", f);
                Assert.Contains(expected, f);
            }
        }

        [Fact]
        public void OldDefaultFalseFields_DoNotCreateClauses()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
            {
                HasAttachment = null,
                IsUnread = null,
                IsFlagged = null,
                Importance = null,
                AttachmentFilter = "any",
                ReadStatus = "any",
                FlagStatus = "any",
                ImportanceFilter = "any",
            });
            Assert.Null(f);
        }

        [Fact]
        public void FirstEmailEverArgs_DoNotContainDefaultFilterClauses()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
            {
                Scope = "all_mail",
                SortOrder = "oldest",
                MaxResults = 1,
                AttachmentFilter = "any",
                ReadStatus = "any",
                FlagStatus = "any",
                ImportanceFilter = "any",
            });

            Assert.Null(f);
        }
    }
}
