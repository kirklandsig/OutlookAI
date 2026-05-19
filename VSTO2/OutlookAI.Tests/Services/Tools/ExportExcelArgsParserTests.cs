using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;
using System.Linq;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class ExportExcelArgsParserTests
    {
        [Fact]
        public void Parse_ValidArgs_ReturnsColumnsAndRows()
        {
            var args = ExportExcelArgsParser.Parse("{"
                + "\"filename_hint\":\" Sales Report \","
                + "\"sheet_name\":\" Export \","
                + "\"columns\":["
                + "{\"name\":\" Date \",\"type\":\"date\"},"
                + "{\"name\":\"Subject\",\"type\":\"text\"},"
                + "{\"name\":\"Amount\",\"type\":\"currency\"},"
                + "{\"name\":\"Reviewed\",\"type\":\"boolean\"}],"
                + "\"rows\":[[\"2026-05-18\",\"Q4\",125.5,true],[\"2026-05-19\",\"Q1\",42,false]]}");

            Assert.Equal("Sales Report", args.FilenameHint);
            Assert.Equal("Export", args.SheetName);
            Assert.Equal(4, args.Columns.Count);
            Assert.Equal("Date", args.Columns[0].Name);
            Assert.Equal(ExcelColumnType.Date, args.Columns[0].Type);
            Assert.Equal(ExcelColumnType.Currency, args.Columns[2].Type);
            Assert.Equal(2, args.Rows.Count);
            Assert.Equal("Q4", args.Rows[0][1].Value<string>());
        }

        [Fact]
        public void Parse_MissingSheetName_FallsBackToFilenameHint()
        {
            var args = ExportExcelArgsParser.Parse("{\"filename_hint\":\" Inbox \",\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}]}");

            Assert.Equal("Inbox", args.FilenameHint);
            Assert.Equal("Inbox", args.SheetName);
        }

        [Fact]
        public void Parse_SheetName_IsTrimmed()
        {
            var args = ExportExcelArgsParser.Parse("{\"filename_hint\":\"Inbox\",\"sheet_name\":\" Sheet 1 \",\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}]}");

            Assert.Equal("Sheet 1", args.SheetName);
        }

        [Fact]
        public void Parse_MissingColumns_ThrowsInvalidArgs()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse("{}"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("columns", ex.Message);
        }

        [Fact]
        public void Parse_EmptyColumns_ThrowsInvalidArgs()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse("{\"columns\":[]}"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("columns", ex.Message);
        }

        [Fact]
        public void Parse_UnknownColumnType_ThrowsInvalidArgsWithType()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse("{\"columns\":[{\"name\":\"Amount\",\"type\":\"money\"}]}"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("type 'money'", ex.Message);
        }

        [Fact]
        public void Parse_MissingColumnName_ThrowsInvalidArgsWithName()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse("{\"columns\":[{\"type\":\"text\"}]}"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("name", ex.Message);
        }

        [Fact]
        public void Parse_TooManyRows_ThrowsTooManyRows()
        {
            var json = "{\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}],\"rows\":[" + string.Join(",", Enumerable.Repeat("[\"x\"]", 10001)) + "]}";

            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse(json));

            Assert.Equal("too_many_rows", ex.Code);
            Assert.Contains("rows", ex.Message);
        }

        [Fact]
        public void Parse_RowShapeMismatch_ThrowsWithRowIndex()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse("{\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"},{\"name\":\"Amount\",\"type\":\"number\"}],\"rows\":[[\"Only one\"]]}"));

            Assert.Equal("row_shape_mismatch", ex.Code);
            Assert.Contains("row 0", ex.Message);
        }

        [Fact]
        public void Parse_EmptyFilenameHint_FallsBackToDefault()
        {
            var args = ExportExcelArgsParser.Parse("{\"filename_hint\":\"   \",\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}]}");

            Assert.Equal("OutlookAI-Report", args.FilenameHint);
            Assert.Equal("OutlookAI-Report", args.SheetName);
        }

        [Fact]
        public void Parse_MissingFilenameHint_FallsBackToDefault()
        {
            var args = ExportExcelArgsParser.Parse("{\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}]}");

            Assert.Equal("OutlookAI-Report", args.FilenameHint);
            Assert.Equal("OutlookAI-Report", args.SheetName);
        }

        [Fact]
        public void Parse_RowsCanBeEmpty()
        {
            var args = ExportExcelArgsParser.Parse("{\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}],\"rows\":[]}");

            Assert.Empty(args.Rows);
        }

        [Fact]
        public void Parse_RowsAbsent_ReturnsEmptyRows()
        {
            var args = ExportExcelArgsParser.Parse("{\"columns\":[{\"name\":\"Subject\",\"type\":\"text\"}]}");

            Assert.Empty(args.Rows);
        }

        [Fact]
        public void Parse_RowCells_ArePreservedAsJTokens()
        {
            var args = ExportExcelArgsParser.Parse("{\"columns\":[{\"name\":\"Count\",\"type\":\"number\"}],\"rows\":[[42]]}");

            Assert.Equal(JTokenType.Integer, args.Rows[0][0].Type);
            Assert.Equal(42, args.Rows[0][0].Value<int>());
        }

        [Fact]
        public void Parse_MalformedJson_ThrowsInvalidArgs()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportExcelArgsParser.Parse("{\"columns\":"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("JSON", ex.Message);
        }
    }
}
