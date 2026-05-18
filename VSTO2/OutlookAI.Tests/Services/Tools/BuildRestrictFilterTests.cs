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
        public void BuildRestrictFilter_PropertyNamesAreQuoted_ForGetTableCompatibility()
        {
            // folder.GetTable() requires DASL property names to be wrapped
            // in double quotes; Items.Restrict() is forgiving and accepts
            // unquoted names. Phase 3b switched the iterative fallback to
            // folder.GetTable() (CollectFolderInputs + AccumulateFolderBuckets),
            // which silently matched zero rows when the filter contained
            // unquoted property URNs like
            //   urn:schemas:httpmail:subject LIKE '%X%'.
            // The user's IT Creations smoke surfaced this: Outlook's UI
            // search found 75 matching emails, every one of our DASL
            // searches returned 0. Quoted property names work for BOTH
            // Restrict and GetTable, so the fix is universal.
            var f = LiveOutlookSurface.BuildRestrictFilter(
                new SearchMessagesArgs { SubjectContains = "IT Creations" });
            Assert.Contains("\"urn:schemas:httpmail:subject\"", f);
            Assert.Contains("LIKE '%IT Creations%'", f);
        }

        [Fact]
        public void QueryOnly_BuildsSubjectOrBodyLike()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { Query = "Q4" });
            Assert.StartsWith("@SQL=", f);
            Assert.Contains("\"urn:schemas:httpmail:subject\" LIKE '%Q4%'", f);
            Assert.Contains("\"urn:schemas:httpmail:textdescription\" LIKE '%Q4%'", f);
        }

        [Fact]
        public void From_MatchesFromHeaderEmailAndSmtp_NoTransportHeadersBlob()
        {
            // The 'from' clause matches three sender-related properties:
            //   urn:schemas:httpmail:from        (RFC 822 From header: "Name" <addr>)
            //   urn:schemas:httpmail:fromemail   (PR_SENDER_EMAIL_ADDRESS; X500 DN for Exchange)
            //   proptag 0x5D01001F               (PR_SENDER_SMTP_ADDRESS; not always populated)
            //
            // Real smoke evidence (2026-05-18): user had 75 UI-search hits
            // for "itcreations" but our from= filter returned 0 across all
            // 200 folders. Root cause: we previously used 'fromname' which
            // is NOT a valid HTTPMAIL URN (canonical is 'from'). The strict
            // folder.GetTable DASL parser silently evaluated that clause as
            // FALSE, then the X500-formatted fromemail and empty SMTP
            // tag killed the other two legs.
            //
            // We intentionally do NOT include PR_TRANSPORT_MESSAGE_HEADERS
            // (proptag 0x007D001F) in this filter even though it always
            // contains the SMTP From line. That property is a multi-KB
            // raw blob, server-NOT-indexed; LIKE-filtering it forces a full
            // per-message read across every folder, which froze Outlook
            // for 10+ minutes on a real mailbox and tripped "trouble
            // connecting to server". The 'from' URN above already exposes
            // the resolved SMTP segment.
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { From = "jane" });
            Assert.Contains("\"urn:schemas:httpmail:from\" LIKE '%jane%'", f);
            Assert.Contains("\"urn:schemas:httpmail:fromemail\" LIKE '%jane%'", f);
            Assert.Contains("\"http://schemas.microsoft.com/mapi/proptag/0x5D01001F\" LIKE '%jane%'", f);
            // The OLD bogus URN must be gone.
            Assert.DoesNotContain("\"urn:schemas:httpmail:fromname\"", f);
            // And the expensive headers blob must NOT be in the filter.
            Assert.DoesNotContain("0x007D001F", f);
        }

        [Fact]
        public void SubjectContains_AddsSubjectLikeClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { SubjectContains = "plan" });
            Assert.Contains("\"urn:schemas:httpmail:subject\" LIKE '%plan%'", f);
            Assert.DoesNotContain("textdescription", f);
        }

        [Fact]
        public void BodyContains_AddsBodyLikeClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { BodyContains = "draft" });
            Assert.Contains("\"urn:schemas:httpmail:textdescription\" LIKE '%draft%'", f);
            Assert.DoesNotContain("subject\" LIKE", f);
        }

        [Fact]
        public void HasAttachment_TrueMapsToAttachmentClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { HasAttachment = true });
            Assert.Contains("\"urn:schemas:httpmail:hasattachment\" = 1", f);
        }

        [Fact]
        public void IsUnread_TrueMapsToUnreadClause()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs { IsUnread = true });
            Assert.Contains("\"urn:schemas:httpmail:read\" = 0", f);
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
        public void DateFrom_UsesIso8601WithTSeparator()
        {
            // Outlook DASL date literals MUST use ISO 8601 with T separator
            // (or US m/d/yyyy format) - a space-separated 'yyyy-MM-dd HH:mm'
            // is silently rejected by folder.GetTable's strict parser, which
            // makes the entire AND-combined filter evaluate FALSE and return
            // 0 rows for every folder. Real smoke evidence (2026-05-18):
            // the model started padding every search with
            //   date_from:"1900-01-01T00:00:00Z", date_to:"<today>T00:00:00Z"
            // and our format string "yyyy-MM-dd HH:mm" produced a bad literal
            // -> every search returned 0 across all 200 folders even though
            // the data was demonstrably there (Outlook UI search found it).
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
            {
                DateFrom = new DateTimeOffset(2026, 5, 10, 14, 30, 0, TimeSpan.Zero),
            });
            // The literal must contain a 'T' between date and time.
            Assert.Matches(@"datereceived"" >= '\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}'", f);
            // And the OLD broken format (space instead of T) must be gone.
            Assert.DoesNotMatch(@"datereceived"" >= '\d{4}-\d{2}-\d{2} \d{2}:\d{2}'", f);
        }

        [Fact]
        public void DateTo_UsesIso8601WithTSeparator()
        {
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
            {
                DateTo = new DateTimeOffset(2026, 5, 18, 23, 59, 59, TimeSpan.Zero),
            });
            Assert.Matches(@"datereceived"" <= '\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}'", f);
            Assert.DoesNotMatch(@"datereceived"" <= '\d{4}-\d{2}-\d{2} \d{2}:\d{2}'", f);
        }

        [Fact]
        public void DateRange_BothBoundsPresent_UsesIso8601()
        {
            // Reproduces the exact model-generated args from the smoke trace:
            //   date_from=1900-01-01T00:00:00Z, date_to=2026-05-18T00:00:00Z
            // This is the case that broke every search.
            var f = LiveOutlookSurface.BuildRestrictFilter(new SearchMessagesArgs
            {
                From = "itcreations",
                DateFrom = new DateTimeOffset(1900, 1, 1, 0, 0, 0, TimeSpan.Zero),
                DateTo = new DateTimeOffset(2026, 5, 18, 0, 0, 0, TimeSpan.Zero),
            });
            Assert.Contains("'1900-01-01T00:00:00'", f);
            Assert.Contains("'2026-05-18T00:00:00'", f);
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
        [InlineData("with", "\"urn:schemas:httpmail:hasattachment\" = 1")]
        [InlineData("without", "\"urn:schemas:httpmail:hasattachment\" = 0")]
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
        [InlineData("unread", "\"urn:schemas:httpmail:read\" = 0")]
        [InlineData("read", "\"urn:schemas:httpmail:read\" = 1")]
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
