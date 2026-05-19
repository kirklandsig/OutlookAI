using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class SearchMessagesArgsParserTests
    {
        [Fact]
        public void ParseSearch_DefaultLookingOldBooleans_DoNotPolluteFilters()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"query\":\"\","
                + "\"from\":\"\","
                + "\"subject_contains\":\"\","
                + "\"body_contains\":\"\","
                + "\"has_attachment\":false,"
                + "\"is_unread\":false,"
                + "\"is_flagged\":false,"
                + "\"importance\":\"normal\"}");

            Assert.Null(args.Query);
            Assert.Null(args.From);
            Assert.Null(args.SubjectContains);
            Assert.Null(args.BodyContains);
            Assert.Null(args.HasAttachment);
            Assert.Null(args.IsUnread);
            Assert.Null(args.IsFlagged);
            Assert.Null(args.Importance);
            Assert.Equal("auto", args.Scope);
            Assert.Equal("newest", args.SortOrder);
            Assert.Equal("any", args.AttachmentFilter);
            Assert.Equal("any", args.ReadStatus);
            Assert.Equal("any", args.FlagStatus);
            Assert.Equal("any", args.ImportanceFilter);
        }

        [Fact]
        public void ParseSearch_NewTriStateFields_AreNormalized()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"query\":\" EIN \","
                + "\"to\":\" Susan \","
                + "\"scope\":\"ALL_MAIL\","
                + "\"sort_order\":\"OLDEST\","
                + "\"attachment_filter\":\"WITH\","
                + "\"read_status\":\"UNREAD\","
                + "\"flag_status\":\"FLAGGED\","
                + "\"importance_filter\":\"HIGH\","
                + "\"date_to\":\"2020-01-01T00:00:00Z\","
                + "\"max_results\":999}");

            Assert.Equal("EIN", args.Query);
            Assert.Equal("Susan", args.To);
            Assert.Equal("all_mail", args.Scope);
            Assert.Equal("oldest", args.SortOrder);
            Assert.Equal("with", args.AttachmentFilter);
            Assert.Equal("unread", args.ReadStatus);
            Assert.Equal("flagged", args.FlagStatus);
            Assert.Equal("high", args.ImportanceFilter);
            Assert.Equal(new DateTimeOffset(2020, 1, 1, 0, 0, 0, TimeSpan.Zero), args.DateTo);
            Assert.Equal(100, args.MaxResults);
        }

        [Fact]
        public void ParseSearch_RecipientAlias_PopulatesToWhenToIsMissing()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{\"recipient\":\"Rich\"}");

            Assert.Equal("Rich", args.To);
        }

        [Fact]
        public void ParseSearch_OldBooleanTrueValues_RemainExplicitFilters()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"has_attachment\":true,"
                + "\"is_unread\":true,"
                + "\"is_flagged\":true,"
                + "\"importance\":\"high\"}");

            Assert.Equal(true, args.HasAttachment);
            Assert.Equal(true, args.IsUnread);
            Assert.Equal(true, args.IsFlagged);
            Assert.Equal("high", args.Importance);
        }

        [Fact]
        public void ParseSearch_NewExplicitNegativeTriStates_AreKept()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"attachment_filter\":\"without\","
                + "\"read_status\":\"read\","
                + "\"flag_status\":\"unflagged\","
                + "\"importance_filter\":\"normal\"}");

            Assert.Equal("without", args.AttachmentFilter);
            Assert.Equal("read", args.ReadStatus);
            Assert.Equal("unflagged", args.FlagStatus);
            Assert.Equal("normal", args.ImportanceFilter);
        }

        [Fact]
        public void ParseCount_DefaultsToIntMaxForMaxResults()
        {
            var args = SearchMessagesArgsParser.ParseCount("{\"scope\":\"all_mail\"}");

            Assert.Equal("all_mail", args.Scope);
            Assert.Equal(int.MaxValue, args.MaxResults);
        }

        [Fact]
        public void ParseSearch_DropsAllTimeSentinelDateBounds()
        {
            var args = SearchMessagesArgsParser.ParseSearch("{"
                + "\"scope\":\"all_mail\","
                + "\"sort_order\":\"oldest\","
                + "\"date_from\":\"1970-01-01T00:00:00Z\","
                + "\"date_to\":\"9999-12-31T00:00:00Z\","
                + "\"max_results\":1}");

            Assert.Null(args.DateFrom);
            Assert.Null(args.DateTo);
            Assert.Equal("all_mail", args.Scope);
            Assert.Equal("oldest", args.SortOrder);
            Assert.Equal(1, args.MaxResults);
        }
    }
}
