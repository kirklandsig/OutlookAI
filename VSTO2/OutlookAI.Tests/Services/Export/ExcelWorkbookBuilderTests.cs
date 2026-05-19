using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using OutlookAI.Services.Export;
using Xunit;

namespace OutlookAI.Tests.Services.Export
{
    public class ExcelWorkbookBuilderTests
    {
        [Fact]
        public void Build_AppliesSheetName()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                Assert.Equal("Export", workbook.Worksheet(1).Name);
            }
        }

        [Fact]
        public void Build_BlankSheetNameFallsBackToSheet1()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("  ", SampleColumns(), SampleRows()))
            {
                Assert.Equal("Sheet1", workbook.Worksheet(1).Name);
            }
        }

        [Fact]
        public void Build_HeaderRowMatchesColumnNames()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                var worksheet = workbook.Worksheet(1);

                Assert.Equal("Date", worksheet.Cell(1, 1).GetString());
                Assert.Equal("Subject", worksheet.Cell(1, 2).GetString());
                Assert.Equal("Sender", worksheet.Cell(1, 3).GetString());
                Assert.Equal("Quoted Total", worksheet.Cell(1, 4).GetString());
                Assert.Equal("Sent At", worksheet.Cell(1, 5).GetString());
            }
        }

        [Fact]
        public void Build_NullHeaderNameRendersEmpty()
        {
            var columns = new List<ExcelColumnSpec>
            {
                new ExcelColumnSpec { Name = null, Type = ExcelColumnType.Text }
            };

            using (var workbook = ExcelWorkbookBuilder.Build("Export", columns, new List<object[]>()))
            {
                Assert.True(workbook.Worksheet(1).Cell(1, 1).IsEmpty());
            }
        }

        [Fact]
        public void Build_HeaderRowIsBoldAndFilled()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                var header = workbook.Worksheet(1).Range(1, 1, 1, 5);

                Assert.All(header.Cells(), cell => Assert.True(cell.Style.Font.Bold));
                Assert.All(header.Cells(), cell => Assert.NotEqual(XLColor.NoColor, cell.Style.Fill.BackgroundColor));
            }
        }

        [Fact]
        public void Build_FreezesTopRow()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                Assert.Equal(1, workbook.Worksheet(1).SheetView.SplitRow);
            }
        }

        [Fact]
        public void Build_AppliesAutofilterOverHeaderAndDataRange()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                var filterRange = workbook.Worksheet(1).AutoFilter.Range;

                Assert.Equal("A1:E3", filterRange.RangeAddress.ToStringRelative());
            }
        }

        [Fact]
        public void Build_DateColumnUsesDateFormat()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                Assert.Equal("yyyy-mm-dd", workbook.Worksheet(1).Cell(2, 1).Style.DateFormat.Format);
            }
        }

        [Fact]
        public void Build_DateTimeColumnUsesDateTimeFormat()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                Assert.Equal("yyyy-mm-dd hh:mm", workbook.Worksheet(1).Cell(2, 5).Style.DateFormat.Format);
            }
        }

        [Fact]
        public void Build_CurrencyColumnUsesDollarFormat()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                Assert.Equal("$#,##0.00", workbook.Worksheet(1).Cell(2, 4).Style.NumberFormat.Format);
            }
        }

        [Fact]
        public void Build_RowCountMatchesHeaderAndInputRows()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), SampleRows()))
            {
                Assert.Equal(3, workbook.Worksheet(1).LastRowUsed().RowNumber());
            }
        }

        [Fact]
        public void Build_EmptyRowsLeavesHeaderOnly()
        {
            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), new List<object[]>()))
            {
                Assert.Equal(1, workbook.Worksheet(1).LastRowUsed().RowNumber());
            }
        }

        [Fact]
        public void Build_NullCellRendersEmpty()
        {
            var rows = new List<object[]>
            {
                new object[] { new DateTime(2026, 5, 18), null, "sender@example.com", 12.5, new DateTime(2026, 5, 18, 14, 21, 0) }
            };

            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), rows))
            {
                Assert.True(workbook.Worksheet(1).Cell(2, 2).IsEmpty());
            }
        }

        [Fact]
        public void Build_ShortRowsLeaveMissingCellsEmptyAndExtraCellsAreIgnored()
        {
            var rows = new List<object[]>
            {
                new object[] { new DateTime(2026, 5, 18), "Short" },
                new object[] { new DateTime(2026, 5, 19), "Extra", "sender@example.com", 99.5, new DateTime(2026, 5, 19, 9, 15, 0), "ignored" }
            };

            using (var workbook = ExcelWorkbookBuilder.Build("Export", SampleColumns(), rows))
            {
                var worksheet = workbook.Worksheet(1);

                Assert.True(worksheet.Cell(2, 3).IsEmpty());
                Assert.True(worksheet.Cell(2, 5).IsEmpty());
                Assert.True(worksheet.Cell(3, 6).IsEmpty());
            }
        }

        [Fact]
        public void Build_NullColumnsThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => ExcelWorkbookBuilder.Build("Export", null, SampleRows()));
        }

        [Fact]
        public void Build_NullRowsThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => ExcelWorkbookBuilder.Build("Export", SampleColumns(), null));
        }

        private static List<ExcelColumnSpec> SampleColumns()
        {
            return new List<ExcelColumnSpec>
            {
                new ExcelColumnSpec { Name = "Date", Type = ExcelColumnType.Date },
                new ExcelColumnSpec { Name = "Subject", Type = ExcelColumnType.Text },
                new ExcelColumnSpec { Name = "Sender", Type = ExcelColumnType.Text },
                new ExcelColumnSpec { Name = "Quoted Total", Type = ExcelColumnType.Currency },
                new ExcelColumnSpec { Name = "Sent At", Type = ExcelColumnType.DateTime }
            };
        }

        private static List<object[]> SampleRows()
        {
            return new List<object[]>
            {
                new object[] { new DateTime(2026, 5, 18), "Budget", "sender@example.com", 12500.0, new DateTime(2026, 5, 18, 14, 21, 0) },
                new object[] { new DateTime(2026, 5, 19), "Follow-up", "other@example.com", 99.5, new DateTime(2026, 5, 19, 9, 15, 0) }
            };
        }
    }
}
