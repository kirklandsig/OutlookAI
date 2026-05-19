using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    internal static class ExportExcelArgsParser
    {
        private const int MaxRows = 10000;
        private const string DefaultFilenameHint = "OutlookAI-Report";

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

            var filenameHint = Clean(args["filename_hint"]) ?? DefaultFilenameHint;
            var sheetName = Clean(args["sheet_name"]) ?? filenameHint;
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

                var name = Clean(column["name"]);
                if (name == null)
                {
                    throw InvalidArgs("columns[" + i + "] name is required");
                }

                var rawType = Clean(column["type"]);
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

        private static string Clean(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            if (token.Type != JTokenType.String) return null;
            var value = token.Value<string>()?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static ToolArgValidationException InvalidArgs(string message)
        {
            return new ToolArgValidationException("invalid_args", message);
        }
    }
}
