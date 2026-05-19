using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExcelColumnTypeTests
    {
        [Theory]
        [InlineData("text", ExcelColumnType.Text)]
        [InlineData("date", ExcelColumnType.Date)]
        [InlineData("datetime", ExcelColumnType.DateTime)]
        [InlineData("number", ExcelColumnType.Number)]
        [InlineData("currency", ExcelColumnType.Currency)]
        [InlineData("boolean", ExcelColumnType.Boolean)]
        [InlineData("TEXT", ExcelColumnType.Text)]
        [InlineData(" date ", ExcelColumnType.Date)]
        public void TryParse_SupportedTypes_ReturnsType(string raw, ExcelColumnType expected)
        {
            var result = ExcelColumnTypeParser.TryParse(raw, out var type);

            Assert.True(result);
            Assert.Equal(expected, type);
        }

        [Theory]
        [InlineData("money")]
        [InlineData("integer")]
        [InlineData("string")]
        [InlineData(null)]
        [InlineData("")]
        public void TryParse_UnsupportedTypes_ReturnsFalse(string raw)
        {
            var result = ExcelColumnTypeParser.TryParse(raw, out var type);

            Assert.False(result);
            Assert.Equal(default(ExcelColumnType), type);
        }
    }
}
