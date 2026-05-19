using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public class ExportPdfArgsParserTests
    {
        [Fact]
        public void Parse_ValidArgs_ReturnsPdfArgs()
        {
            var args = ExportPdfArgsParser.Parse("{"
                + "\"filename_hint\":\" Weekly Report \","
                + "\"content_markdown\":\"# Report\\nBody\","
                + "\"title\":\" Inbox Summary \","
                + "\"subtitle\":\" Last 7 Days \"}");

            Assert.Equal("Weekly Report", args.FilenameHint);
            Assert.Equal("# Report\nBody", args.ContentMarkdown);
            Assert.Equal("Inbox Summary", args.Title);
            Assert.Equal("Last 7 Days", args.Subtitle);
        }

        [Fact]
        public void Parse_MissingContentMarkdown_ThrowsInvalidArgsWithContentMarkdown()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse("{}"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("content_markdown", ex.Message);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Parse_BlankArgs_ThrowsInvalidArgsWithContentMarkdown(string argsJson)
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse(argsJson));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("content_markdown", ex.Message);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        public void Parse_EmptyContentMarkdown_ThrowsInvalidArgs(string content)
        {
            var json = "{\"content_markdown\":\"" + content + "\"}";

            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse(json));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("content_markdown", ex.Message);
        }

        [Fact]
        public void Parse_ContentMarkdownTooLarge_ThrowsContentTooLarge()
        {
            var content = new string('x', 250001);
            var json = "{\"content_markdown\":\"" + content + "\"}";

            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse(json));

            Assert.Equal("content_too_large", ex.Code);
            Assert.Contains("content_markdown", ex.Message);
        }

        [Fact]
        public void Parse_TitleLongerThan200Characters_IsClamped()
        {
            var title = new string('t', 201);
            var args = ExportPdfArgsParser.Parse("{\"content_markdown\":\"body\",\"title\":\"" + title + "\"}");

            Assert.Equal(200, args.Title.Length);
            Assert.Equal(new string('t', 200), args.Title);
        }

        [Fact]
        public void Parse_SubtitleLongerThan400Characters_IsClamped()
        {
            var subtitle = new string('s', 401);
            var args = ExportPdfArgsParser.Parse("{\"content_markdown\":\"body\",\"subtitle\":\"" + subtitle + "\"}");

            Assert.Equal(400, args.Subtitle.Length);
            Assert.Equal(new string('s', 400), args.Subtitle);
        }

        [Fact]
        public void Parse_MissingFilenameHint_FallsBackToDefault()
        {
            var args = ExportPdfArgsParser.Parse("{\"content_markdown\":\"body\"}");

            Assert.Equal("OutlookAI-Report", args.FilenameHint);
        }

        [Fact]
        public void Parse_MissingFilenameHintWithTitle_FallsBackToTitle()
        {
            var args = ExportPdfArgsParser.Parse("{\"content_markdown\":\"body\",\"title\":\" Report Title \"}");

            Assert.Equal("Report Title", args.FilenameHint);
        }

        [Fact]
        public void Parse_MissingTitleAndSubtitle_DefaultToNull()
        {
            var args = ExportPdfArgsParser.Parse("{\"content_markdown\":\"body\"}");

            Assert.Null(args.Title);
            Assert.Null(args.Subtitle);
        }

        [Fact]
        public void Parse_MalformedJson_ThrowsInvalidArgs()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse("{\"content_markdown\":"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("JSON", ex.Message);
        }

        [Fact]
        public void Parse_ContentMarkdownNonString_ThrowsInvalidArgsWithContentMarkdown()
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse("{\"content_markdown\":123}"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains("content_markdown", ex.Message);
        }

        [Theory]
        [InlineData("filename_hint")]
        [InlineData("title")]
        [InlineData("subtitle")]
        public void Parse_OptionalStringNonString_ThrowsInvalidArgsWithField(string fieldName)
        {
            var ex = Assert.Throws<ToolArgValidationException>(() => ExportPdfArgsParser.Parse("{\"content_markdown\":\"body\",\"" + fieldName + "\":{}}"));

            Assert.Equal("invalid_args", ex.Code);
            Assert.Contains(fieldName, ex.Message);
        }

        [Fact]
        public void Parse_ContentMarkdown_PreservesWhitespace()
        {
            var args = ExportPdfArgsParser.Parse("{\"content_markdown\":\"  body  \",\"title\":\"Title\"}");

            Assert.Equal("  body  ", args.ContentMarkdown);
        }
    }
}
