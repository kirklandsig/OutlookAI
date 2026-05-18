using System;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class AggregateMessagesArgsParserTests
    {
        [Fact]
        public void Parse_EmptyJson_AppliesDefaults()
        {
            var args = AggregateMessagesArgsParser.Parse("{}");
            Assert.Equal("auto", args.Scope);
            Assert.Equal("sender", args.GroupBy);
            Assert.Equal(10, args.TopN);
            Assert.Null(args.From);
            Assert.Null(args.SubjectContains);
            Assert.Null(args.BodyContains);
            Assert.Null(args.DateFrom);
            Assert.Null(args.DateTo);
        }

        [Fact]
        public void Parse_AllFields_RoundTrips()
        {
            var json = "{"
                + "\"scope\":\"all_mail\","
                + "\"folder_id\":\"f1\","
                + "\"date_from\":\"2026-05-01T00:00:00Z\","
                + "\"date_to\":\"2026-05-31T00:00:00Z\","
                + "\"from\":\"jane@example.com\","
                + "\"subject_contains\":\"Q4\","
                + "\"body_contains\":\"draft\","
                + "\"group_by\":\"day\","
                + "\"top_n\":25}";
            var args = AggregateMessagesArgsParser.Parse(json);
            Assert.Equal("all_mail", args.Scope);
            Assert.Equal("f1", args.FolderId);
            Assert.Equal(new DateTimeOffset(2026, 5, 1, 0, 0, 0, TimeSpan.Zero), args.DateFrom);
            Assert.Equal(new DateTimeOffset(2026, 5, 31, 0, 0, 0, TimeSpan.Zero), args.DateTo);
            Assert.Equal("jane@example.com", args.From);
            Assert.Equal("Q4", args.SubjectContains);
            Assert.Equal("draft", args.BodyContains);
            Assert.Equal("day", args.GroupBy);
            Assert.Equal(25, args.TopN);
        }

        [Theory]
        [InlineData("SENDER", "sender")]
        [InlineData("Day", "day")]
        [InlineData("folder", "folder")]
        [InlineData("garbage", "sender")]    // unknown -> default "sender"
        [InlineData("", "sender")]           // empty -> default
        public void Parse_GroupBy_NormalizedOrDefaulted(string raw, string expected)
        {
            var args = AggregateMessagesArgsParser.Parse("{\"group_by\":\"" + raw + "\"}");
            Assert.Equal(expected, args.GroupBy);
        }

        [Theory]
        [InlineData(0, 1)]                   // floored to 1
        [InlineData(-5, 1)]                  // floored to 1
        [InlineData(50, 50)]
        [InlineData(9999, 100)]              // capped at 100
        public void Parse_TopN_ClampedToValidRange(int input, int expected)
        {
            var args = AggregateMessagesArgsParser.Parse("{\"top_n\":" + input + "}");
            Assert.Equal(expected, args.TopN);
        }

        [Theory]
        [InlineData("ALL_MAIL", "all_mail")]
        [InlineData("Current_Folder", "current_folder")]
        [InlineData("Auto", "auto")]
        [InlineData("garbage", "auto")]
        public void Parse_Scope_NormalizedOrDefaulted(string raw, string expected)
        {
            var args = AggregateMessagesArgsParser.Parse("{\"scope\":\"" + raw + "\"}");
            Assert.Equal(expected, args.Scope);
        }

        [Fact]
        public void Parse_NullOrWhitespaceJson_AppliesDefaults()
        {
            Assert.Equal("auto", AggregateMessagesArgsParser.Parse(null).Scope);
            Assert.Equal("auto", AggregateMessagesArgsParser.Parse("").Scope);
            Assert.Equal("auto", AggregateMessagesArgsParser.Parse("   ").Scope);
        }
    }
}
