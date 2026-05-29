using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    internal static class ExportExcelArgsParser
    {
        private const int MaxRows = BulkExportRowCap.Max;
        private const string DefaultFilenameHint = "OutlookAI-Report";
        private static readonly char[] InvalidSheetNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

        public static ExportExcelArgs Parse(string argsJson)
        {
            JObject args;
            try
            {
                args = JObject.Parse(string.IsNullOrWhiteSpace(argsJson) ? "{}" : argsJson);
            }
            catch (JsonException ex)
            {
                throw InvalidArgs("Invalid JSON args: " + ex.Message);
            }

            var filenameHint = CleanOptionalString(args["filename_hint"], "filename_hint") ?? DefaultFilenameHint;
            var sheetName = CleanOptionalString(args["sheet_name"], "sheet_name") ?? filenameHint;
            ValidateSheetName(sheetName);
            var columns = ParseColumns(args["columns"]);

            return new ExportExcelArgs
            {
                FilenameHint = filenameHint,
                SheetName = sheetName,
                Columns = columns,
                Rows = ParseRows(args["rows"], columns.Count),
            };
        }

        private static IList<ExcelColumnSpec> ParseColumns(JToken token)
        {
            if (token == null)
            {
                throw InvalidArgs("columns is required");
            }

            var columnArray = token as JArray;
            if (columnArray == null)
            {
                throw InvalidArgs("columns must be an array");
            }

            if (columnArray.Count == 0)
            {
                throw InvalidArgs("columns must not be empty");
            }

            var columns = new List<ExcelColumnSpec>();
            for (var i = 0; i < columnArray.Count; i++)
            {
                var column = columnArray[i] as JObject;
                if (column == null)
                {
                    throw InvalidArgs("columns[" + i + "] must be an object");
                }

                var name = CleanRequiredString(column["name"], "columns[" + i + "] name");
                var rawType = CleanRequiredString(column["type"], "columns[" + i + "] type");
                ExcelColumnType type;
                if (!ExcelColumnTypeParser.TryParse(rawType, out type))
                {
                    throw InvalidArgs("columns[" + i + "] has invalid type '" + (rawType ?? "") + "'");
                }

                columns.Add(new ExcelColumnSpec { Name = name, Type = type });
            }

            return columns;
        }

        private static IList<JToken[]> ParseRows(JToken token, int columnCount)
        {
            var rows = new List<JToken[]>();
            if (token == null || token.Type == JTokenType.Null)
            {
                return rows;
            }

            var rowArray = token as JArray;
            if (rowArray == null)
            {
                throw InvalidArgs("rows must be an array");
            }

            if (rowArray.Count > MaxRows)
            {
                throw new ToolArgValidationException("too_many_rows", "rows count must be <= " + MaxRows);
            }

            for (var i = 0; i < rowArray.Count; i++)
            {
                var row = rowArray[i] as JArray;
                if (row == null)
                {
                    throw InvalidArgs("row " + i + " must be an array");
                }

                if (row.Count != columnCount)
                {
                    throw new ToolArgValidationException("row_shape_mismatch", "row " + i + " has " + row.Count + " cells but columns has " + columnCount);
                }

                var cells = new JToken[row.Count];
                for (var j = 0; j < row.Count; j++)
                {
                    cells[j] = row[j];
                }
                rows.Add(cells);
            }

            return rows;
        }

        private static string CleanOptionalString(JToken token, string fieldName)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            if (token.Type != JTokenType.String) throw InvalidArgs(fieldName + " must be a string");
            var value = token.Value<string>()?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static string CleanRequiredString(JToken token, string fieldName)
        {
            if (token == null || token.Type == JTokenType.Null) throw InvalidArgs(fieldName + " is required");
            if (token.Type != JTokenType.String) throw InvalidArgs(fieldName + " must be a string");
            var value = token.Value<string>()?.Trim();
            if (string.IsNullOrEmpty(value)) throw InvalidArgs(fieldName + " is required");
            return value;
        }

        private static void ValidateSheetName(string sheetName)
        {
            if (sheetName.Length > 31)
            {
                throw InvalidArgs("sheet_name must be 31 characters or fewer");
            }

            if (sheetName.IndexOfAny(InvalidSheetNameChars) >= 0)
            {
                throw InvalidArgs("sheet_name contains an invalid Excel worksheet character");
            }

            if (string.Equals(sheetName, "History", StringComparison.OrdinalIgnoreCase))
            {
                throw InvalidArgs("sheet_name 'History' is reserved by Excel");
            }
        }

        private static ToolArgValidationException InvalidArgs(string message)
        {
            return new ToolArgValidationException("invalid_args", message);
        }
    }
}
