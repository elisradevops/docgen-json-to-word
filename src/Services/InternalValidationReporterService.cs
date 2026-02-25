using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class InternalValidationReporterService : IInternalValidationReporterService
    {
        private readonly ILogger<InternalValidationReporterService> _logger;
        private readonly ISpreadsheetService _spreadsheetService;
        private readonly IStylesheetService _stylesheetService;

        private static readonly string[] DefaultColumnOrder = new[]
        {
            "Test Case ID",
            "Test Case Title",
            "Mentioned but Not Linked",
            "Linked but Not Mentioned",
            "Validation Status",
        };

        public InternalValidationReporterService(
            ILogger<InternalValidationReporterService> logger,
            ISpreadsheetService spreadsheetService,
            IStylesheetService stylesheetService
        )
        {
            _logger = logger;
            _spreadsheetService = spreadsheetService;
            _stylesheetService = stylesheetService;
        }

        public void Insert(
            SpreadsheetDocument document,
            string worksheetName,
            InternalValidationReporterModel coverageModel
        )
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (coverageModel == null)
                throw new ArgumentNullException(nameof(coverageModel));

            var safeName = string.IsNullOrWhiteSpace(worksheetName)
                ? "MEWP Internal Validation"
                : worksheetName;

            try
            {
                WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
                if (workbookPart.Workbook == null)
                {
                    workbookPart.Workbook = new Workbook();
                }

                WorksheetPart worksheetPart = _spreadsheetService.GetOrCreateWorksheetPart(workbookPart, safeName);
                worksheetPart.Worksheet = new Worksheet();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet.Append(sheetData);
                _stylesheetService.EnsureStylesheet(workbookPart);

                SheetViews sheetViews = new SheetViews(
                    new SheetView { WorkbookViewId = 0U, RightToLeft = false }
                );
                worksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);

                var columnOrder =
                    coverageModel.ColumnOrder != null && coverageModel.ColumnOrder.Count > 0
                        ? coverageModel.ColumnOrder.ToArray()
                        : DefaultColumnOrder;

                var columns = CreateColumns(columnOrder);
                worksheetPart.Worksheet.InsertBefore(columns, sheetData);

                uint rowIndex = 1;
                Row headerRow = new Row { RowIndex = rowIndex };
                sheetData.Append(headerRow);
                for (int i = 0; i < columnOrder.Length; i++)
                {
                    string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                    Cell headerCell = _spreadsheetService.CreateTextCell(
                        $"{columnLetter}{rowIndex}",
                        columnOrder[i],
                        1
                    );
                    headerRow.Append(headerCell);
                }

                var rows = coverageModel.Rows ?? new List<Dictionary<string, object>>();
                foreach (var row in rows)
                {
                    rowIndex++;
                    Row dataRow = new Row { RowIndex = rowIndex };
                    sheetData.Append(dataRow);
                    bool useFirstAlternatingColor = rowIndex % 2 == 0;

                    for (int i = 0; i < columnOrder.Length; i++)
                    {
                        string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                        string cellRef = $"{columnLetter}{rowIndex}";
                        object value = null;
                        if (row != null)
                        {
                            row.TryGetValue(columnOrder[i], out value);
                        }
                        var cellValue = value?.ToString() ?? string.Empty;
                        uint styleIndex = ResolveDataStyle(columnOrder[i], useFirstAlternatingColor);
                        Cell cell;
                        if (IsNumericColumn(columnOrder[i]) && TryToNumberCellValue(value, out string numericValue))
                        {
                            cell = _spreadsheetService.CreateNumberCell(cellRef, numericValue, styleIndex);
                        }
                        else
                        {
                            cell = _spreadsheetService.CreateTextCell(cellRef, cellValue, styleIndex);
                        }
                        dataRow.Append(cell);
                    }
                }

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    $"Error inserting Internal Validation worksheet '{worksheetName}': {ex.Message}"
                );
                throw;
            }
        }

        private Columns CreateColumns(string[] columnOrder)
        {
            Columns columns = new Columns();
            for (int i = 0; i < columnOrder.Length; i++)
            {
                string name = columnOrder[i] ?? string.Empty;
                double width = ResolveColumnWidth(name);
                columns.Append(new Column
                {
                    Min = (uint)(i + 1),
                    Max = (uint)(i + 1),
                    Width = width,
                    CustomWidth = true,
                });
            }
            return columns;
        }

        private static double ResolveColumnWidth(string columnName)
        {
            var key = (columnName ?? string.Empty).Trim().ToLowerInvariant();
            if (key == "test case id") return 16;
            if (key == "test case title") return 34;
            if (key == "mentioned but not linked") return 48;
            if (key == "linked but not mentioned") return 36;
            if (key == "validation status") return 20;
            return 34;
        }

        private static uint ResolveDataStyle(string columnName, bool useFirstAlternatingColor)
        {
            if (IsNumericColumn(columnName))
            {
                return useFirstAlternatingColor ? 10U : 11U;
            }

            return useFirstAlternatingColor ? 6U : 7U;
        }

        private static bool IsNumericColumn(string columnName)
        {
            var key = (columnName ?? string.Empty).Trim().ToLowerInvariant();
            return key == "test case id";
        }

        private static bool TryToNumberCellValue(object value, out string numericValue)
        {
            numericValue = string.Empty;
            if (value == null)
                return false;

            if (value is byte or short or int or long or sbyte or ushort or uint or ulong)
            {
                numericValue = Convert.ToString(value, CultureInfo.InvariantCulture);
                return !string.IsNullOrWhiteSpace(numericValue);
            }

            var text = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal parsedInvariant))
            {
                numericValue = parsedInvariant.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.CurrentCulture, out decimal parsedCurrent))
            {
                numericValue = parsedCurrent.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            return false;
        }
    }
}
