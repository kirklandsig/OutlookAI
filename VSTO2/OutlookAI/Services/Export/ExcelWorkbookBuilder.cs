using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace OutlookAI.Services.Export
{
    public static class ExcelWorkbookBuilder
    {
        public static XLWorkbook Build(string sheetName, IList<ExcelColumnSpec> columns, IList<object[]> rows)
        {
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(string.IsNullOrWhiteSpace(sheetName) ? "Sheet1" : sheetName);

            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var cell = worksheet.Cell(1, columnIndex + 1);
                cell.Value = columns[columnIndex].Name ?? string.Empty;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
            }

            worksheet.SheetView.FreezeRows(1);

            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                if (row == null) continue;

                var maxCellCount = Math.Min(columns.Count, row.Length);
                for (var columnIndex = 0; columnIndex < maxCellCount; columnIndex++)
                {
                    var value = row[columnIndex];
                    if (value == null) continue;

                    var cell = worksheet.Cell(rowIndex + 2, columnIndex + 1);
                    SetCellValue(cell, value);
                    ApplyFormat(cell, columns[columnIndex].Type);
                }
            }

            if (columns.Count > 0)
            {
                worksheet.Range(1, 1, Math.Max(rows.Count + 1, 1), columns.Count).SetAutoFilter();
                worksheet.Columns(1, columns.Count).AdjustToContents();
            }

            return workbook;
        }

        private static void SetCellValue(IXLCell cell, object value)
        {
            if (value is DateTime dateTime)
            {
                cell.Value = dateTime;
                return;
            }

            if (value is bool boolean)
            {
                cell.Value = boolean;
                return;
            }

            if (value is byte || value is sbyte || value is short || value is ushort || value is int || value is uint || value is long || value is ulong || value is float || value is double || value is decimal)
            {
                cell.Value = Convert.ToDouble(value);
                return;
            }

            cell.Value = value.ToString();
        }

        private static void ApplyFormat(IXLCell cell, ExcelColumnType type)
        {
            switch (type)
            {
                case ExcelColumnType.Currency:
                    cell.Style.NumberFormat.Format = "$#,##0.00";
                    break;
                case ExcelColumnType.Date:
                    cell.Style.DateFormat.Format = "yyyy-mm-dd";
                    break;
                case ExcelColumnType.DateTime:
                    cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm";
                    break;
            }
        }
    }
}
