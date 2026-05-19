using System;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Export
{
    public static class ExcelCellCoercer
    {
        public static object Coerce(JToken value, ExcelColumnType type)
        {
            if (IsNull(value)) return null;

            switch (type)
            {
                case ExcelColumnType.Date:
                case ExcelColumnType.DateTime:
                    return CoerceDateTime(value);
                case ExcelColumnType.Number:
                    return CoerceNumber(value);
                case ExcelColumnType.Currency:
                    return CoerceCurrency(value);
                case ExcelColumnType.Boolean:
                    return CoerceBoolean(value);
                default:
                    return ToText(value);
            }
        }

        private static object CoerceDateTime(JToken value)
        {
            if (value.Type == JTokenType.Date) return value.Value<DateTime>();

            var text = ToText(value);
            DateTime parsed;
            return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind | DateTimeStyles.AllowWhiteSpaces, out parsed)
                ? (object)parsed
                : text;
        }

        private static object CoerceNumber(JToken value)
        {
            if (IsNumeric(value)) return Convert.ToDouble(((JValue)value).Value, CultureInfo.InvariantCulture);

            var text = ToText(value);
            double parsed;
            return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out parsed)
                ? (object)parsed
                : text;
        }

        private static object CoerceCurrency(JToken value)
        {
            if (IsNumeric(value)) return Convert.ToDouble(((JValue)value).Value, CultureInfo.InvariantCulture);

            var text = ToText(value);
            var candidate = text.Trim();
            if (candidate.StartsWith("$", StringComparison.Ordinal)) candidate = candidate.Substring(1);
            candidate = candidate.Replace(",", "");

            double parsed;
            return double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)
                ? (object)parsed
                : text;
        }

        private static object CoerceBoolean(JToken value)
        {
            if (value.Type == JTokenType.Boolean) return value.Value<bool>();

            var text = ToText(value);
            bool parsed;
            return bool.TryParse(text, out parsed) ? (object)parsed : text;
        }

        private static string ToText(JToken value)
        {
            if (value.Type == JTokenType.String) return value.Value<string>();
            return value.ToString(Formatting.None);
        }

        private static bool IsNull(JToken value)
        {
            return value == null || value.Type == JTokenType.Null || value.Type == JTokenType.Undefined;
        }

        private static bool IsNumeric(JToken value)
        {
            return value is JValue && (value.Type == JTokenType.Integer || value.Type == JTokenType.Float);
        }
    }
}
