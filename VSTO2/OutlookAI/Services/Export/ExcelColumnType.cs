using System;

namespace OutlookAI.Services.Export
{
    public enum ExcelColumnType
    {
        Text,
        Date,
        DateTime,
        Number,
        Currency,
        Boolean
    }

    public static class ExcelColumnTypeParser
    {
        public static bool TryParse(string raw, out ExcelColumnType type)
        {
            type = default(ExcelColumnType);
            if (string.IsNullOrWhiteSpace(raw)) return false;

            switch (raw.Trim().ToLowerInvariant())
            {
                case "text":
                    type = ExcelColumnType.Text;
                    return true;
                case "date":
                    type = ExcelColumnType.Date;
                    return true;
                case "datetime":
                    type = ExcelColumnType.DateTime;
                    return true;
                case "number":
                    type = ExcelColumnType.Number;
                    return true;
                case "currency":
                    type = ExcelColumnType.Currency;
                    return true;
                case "boolean":
                    type = ExcelColumnType.Boolean;
                    return true;
                default:
                    return false;
            }
        }
    }
}
