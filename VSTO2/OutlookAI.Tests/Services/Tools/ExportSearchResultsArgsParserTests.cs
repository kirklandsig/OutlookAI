using System.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class ExportSearchResultsArgsParserTests
    {
        [Fact]
        public void Parse_ExtractsFilterAndColumns()
        {
            var json = "{\"from\":\"IT Creations\",\"scope\":\"all_mail\",\"columns\":[\"subject\",\"from\",\"received_at\",\"snippet\"],\"filename_hint\":\"vendor\"}";
            var args = ExportSearchResultsArgsParser.Parse(json);

            Assert.Equal("IT Creations", args.Filter.From);
            Assert.Equal("all_mail", args.Filter.Scope);
            Assert.Equal(new[] { "subject", "from", "received_at", "snippet" }, args.Columns.ToArray());
            Assert.Equal("vendor", args.FilenameHint);
        }

        [Fact]
        public void Parse_DropsUnknownColumns()
        {
            var json = "{\"columns\":[\"subject\",\"bogus\",\"snippet\"]}";
            var args = ExportSearchResultsArgsParser.Parse(json);
            Assert.Equal(new[] { "subject", "snippet" }, args.Columns.ToArray());
        }

        [Fact]
        public void Parse_DefaultsColumnsWhenNoneProvided()
        {
            var args = ExportSearchResultsArgsParser.Parse("{\"from\":\"x\"}");
            Assert.Equal(ExportSearchResultsArgsParser.DefaultColumns, args.Columns.ToArray());
        }

        [Fact]
        public void Parse_AllUnknownColumns_FallsBackToDefaults()
        {
            var args = ExportSearchResultsArgsParser.Parse("{\"columns\":[\"bogus\",\"nope\"]}");
            Assert.Equal(ExportSearchResultsArgsParser.DefaultColumns, args.Columns.ToArray());
        }

        [Fact]
        public void Parse_DeduplicatesColumns()
        {
            var args = ExportSearchResultsArgsParser.Parse("{\"columns\":[\"subject\",\"subject\",\"from\"]}");
            Assert.Equal(new[] { "subject", "from" }, args.Columns.ToArray());
        }
    }
}
