using System;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExcelCellCoercerTests
    {
        [Fact]
        public void Coerce_TextString_ReturnsRawString()
        {
            var result = ExcelCellCoercer.Coerce(new JValue("hello"), ExcelColumnType.Text);

            Assert.Equal("hello", result);
        }

        [Fact]
        public void Coerce_TextNumber_ReturnsStringifiedNumber()
        {
            var result = ExcelCellCoercer.Coerce(new JValue(42), ExcelColumnType.Text);

            Assert.Equal("42", result);
        }

        [Fact]
        public void Coerce_DateIsoString_ReturnsDateTime()
        {
            var result = Assert.IsType<DateTime>(ExcelCellCoercer.Coerce(new JValue("2026-05-18"), ExcelColumnType.Date));

            Assert.Equal(new DateTime(2026, 5, 18), result.Date);
        }

        [Fact]
        public void Coerce_DateTimeIsoString_ReturnsDateTime()
        {
            var result = Assert.IsType<DateTime>(ExcelCellCoercer.Coerce(new JValue("2026-05-18T14:21:00Z"), ExcelColumnType.DateTime));

            Assert.Equal(2026, result.Year);
            Assert.Equal(5, result.Month);
            Assert.Equal(18, result.Day);
            Assert.Equal(14, result.Hour);
            Assert.Equal(21, result.Minute);
        }

        [Fact]
        public void Coerce_DateUnparseable_ReturnsOriginalText()
        {
            var result = ExcelCellCoercer.Coerce(new JValue("yesterday"), ExcelColumnType.Date);

            Assert.Equal("yesterday", result);
        }

        [Fact]
        public void Coerce_NumberRawNumeric_ReturnsDouble()
        {
            var result = ExcelCellCoercer.Coerce(new JValue(12500.5), ExcelColumnType.Number);

            Assert.Equal(12500.5, Assert.IsType<double>(result));
        }

        [Fact]
        public void Coerce_NumberNumericString_ReturnsDouble()
        {
            var result = ExcelCellCoercer.Coerce(new JValue("12500"), ExcelColumnType.Number);

            Assert.Equal(12500.0, Assert.IsType<double>(result));
        }

        [Fact]
        public void Coerce_NumberUnparseable_ReturnsOriginalText()
        {
            var result = ExcelCellCoercer.Coerce(new JValue("twelve"), ExcelColumnType.Number);

            Assert.Equal("twelve", result);
        }

        [Fact]
        public void Coerce_CurrencyDollarString_ReturnsDouble()
        {
            var result = ExcelCellCoercer.Coerce(new JValue("$12,500.00"), ExcelColumnType.Currency);

            Assert.Equal(12500.0, Assert.IsType<double>(result));
        }

        [Fact]
        public void Coerce_CurrencyRawNumber_ReturnsDouble()
        {
            var result = ExcelCellCoercer.Coerce(new JValue(12500), ExcelColumnType.Currency);

            Assert.Equal(12500.0, Assert.IsType<double>(result));
        }

        [Fact]
        public void Coerce_CurrencyExoticSymbol_ReturnsOriginalText()
        {
            var result = ExcelCellCoercer.Coerce(new JValue("\u20B91,200"), ExcelColumnType.Currency);

            Assert.Equal("\u20B91,200", result);
        }

        [Theory]
        [InlineData(true, true)]
        [InlineData(false, false)]
        [InlineData("true", true)]
        [InlineData("FALSE", false)]
        public void Coerce_BooleanValues_ReturnsBool(object input, bool expected)
        {
            var result = ExcelCellCoercer.Coerce(new JValue(input), ExcelColumnType.Boolean);

            Assert.Equal(expected, Assert.IsType<bool>(result));
        }

        [Fact]
        public void Coerce_NullJToken_ReturnsNullForTextAndCurrency()
        {
            Assert.Null(ExcelCellCoercer.Coerce(null, ExcelColumnType.Text));
            Assert.Null(ExcelCellCoercer.Coerce(null, ExcelColumnType.Currency));
        }
    }
}
