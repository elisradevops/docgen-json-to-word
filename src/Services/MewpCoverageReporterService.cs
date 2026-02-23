using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class MewpCoverageReporterService : IMewpCoverageReporterService
    {
        private readonly ILogger<MewpCoverageReporterService> _logger;
        private readonly ISpreadsheetService _spreadsheetService;
        private readonly IStylesheetService _stylesheetService;

        private static readonly string[] DefaultColumnOrder = new[]
        {
            "L2 REQ ID",
            "L2 REQ Title",
            "L2 SubSystem",
            "L2 Run Status",
            "Bug ID",
            "Bug Title",
            "Bug Responsibility",
            "L3 REQ ID",
            "L3 REQ Title",
            "L4 REQ ID",
            "L4 REQ Title",
        };
        
        private static readonly string[] RequirementMergeCandidateColumns = new[]
        {
            "L2 REQ ID",
            "L2 REQ Title",
            "L2 SubSystem",
            "L2 Run Status",
            "L3 REQ ID",
            "L3 REQ Title",
        };

        private static readonly HashSet<string> BugColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Bug ID",
            "Bug Title",
            "Bug Responsibility",
        };

        private static readonly HashSet<string> LinkedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "L3 REQ ID",
            "L3 REQ Title",
            "L4 REQ ID",
            "L4 REQ Title",
        };

        public MewpCoverageReporterService(
            ILogger<MewpCoverageReporterService> logger,
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
            MewpCoverageReporterModel coverageModel
        )
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (coverageModel == null)
                throw new ArgumentNullException(nameof(coverageModel));

            var safeName = string.IsNullOrWhiteSpace(worksheetName)
                ? "MEWP L2 Coverage"
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
                MergeCells mergeCells = new MergeCells();

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
                var useFirstAlternatingColorByRow = BuildRowAlternatingColorFlags(
                    rows,
                    columnOrder,
                    coverageModel.MergeDuplicateRequirementCells
                );
                foreach (var row in rows)
                {
                    rowIndex++;
                    Row dataRow = new Row { RowIndex = rowIndex };
                    sheetData.Append(dataRow);
                    int dataRowOffset = (int)rowIndex - 2;
                    bool useFirstAlternatingColor =
                        dataRowOffset >= 0 &&
                        dataRowOffset < useFirstAlternatingColorByRow.Count
                            ? useFirstAlternatingColorByRow[dataRowOffset]
                            : rowIndex % 2 == 0;

                    for (int i = 0; i < columnOrder.Length; i++)
                    {
                        string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                        string cellRef = $"{columnLetter}{rowIndex}";
                        object value = null;
                        string cellValue = string.Empty;
                        if (row != null && row.TryGetValue(columnOrder[i], out value))
                        {
                            cellValue = value?.ToString() ?? string.Empty;
                        }
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

                if (coverageModel.MergeDuplicateRequirementCells)
                {
                    AppendRequirementDuplicateMergeRanges(mergeCells, rows, columnOrder);
                }

                if (mergeCells.Any())
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, sheetData);
                }

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }
            catch (Exception ex)
            {
                _logger.LogError(
                    ex,
                    $"Error inserting MEWP coverage worksheet '{worksheetName}': {ex.Message}"
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
            if (key == "l2 req id") return 18;
            if (key == "l2 req title") return 40;
            if (key == "l2 subsystem") return 24;
            if (key == "l2 run status") return 18;
            if (key == "bug id") return 14;
            if (key == "bug title") return 34;
            if (key == "bug responsibility") return 22;
            if (key == "l3 req id") return 16;
            if (key == "l3 req title") return 30;
            if (key == "l4 req id") return 16;
            if (key == "l4 req title") return 30;
            return 24;
        }

        private static uint ResolveDataStyle(string columnName, bool useFirstAlternatingColor)
        {
            if (IsNumericColumn(columnName))
            {
                if (IsBugColumn(columnName))
                {
                    return useFirstAlternatingColor ? 20U : 21U;
                }
                return useFirstAlternatingColor ? 10U : 11U;
            }

            if (IsBugColumn(columnName))
            {
                return useFirstAlternatingColor ? 18U : 19U;
            }

            if (IsLinkedColumn(columnName))
            {
                return useFirstAlternatingColor ? 22U : 23U;
            }

            return useFirstAlternatingColor ? 6U : 7U;
        }

        private static bool IsNumericColumn(string columnName)
        {
            var key = (columnName ?? string.Empty).Trim().ToLowerInvariant();
            return key == "bug id";
        }

        private static bool IsBugColumn(string columnName)
        {
            return !string.IsNullOrWhiteSpace(columnName) && BugColumns.Contains(columnName.Trim());
        }

        private static bool IsLinkedColumn(string columnName)
        {
            return !string.IsNullOrWhiteSpace(columnName) && LinkedColumns.Contains(columnName.Trim());
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

            if (value is float or double or decimal)
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

        private List<bool> BuildRowAlternatingColorFlags(
            IReadOnlyList<Dictionary<string, object>> rows,
            IReadOnlyList<string> columnOrder,
            bool groupByL2
        )
        {
            var result = new List<bool>();
            if (rows == null || rows.Count == 0)
            {
                return result;
            }

            if (!groupByL2)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    result.Add(i % 2 == 0);
                }
                return result;
            }

            int l2IdColumnIndex = FindColumnIndex(columnOrder, "L2 REQ ID");
            if (l2IdColumnIndex < 0)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    result.Add(i % 2 == 0);
                }
                return result;
            }

            bool useFirstAlternatingColor = true;
            string previousL2Id = null;

            for (int i = 0; i < rows.Count; i++)
            {
                string currentL2Id = GetComparableCellValue(rows[i], columnOrder[l2IdColumnIndex]);
                if (i > 0 && !string.Equals(previousL2Id, currentL2Id, StringComparison.OrdinalIgnoreCase))
                {
                    useFirstAlternatingColor = !useFirstAlternatingColor;
                }

                result.Add(useFirstAlternatingColor);
                previousL2Id = currentL2Id;
            }

            return result;
        }

        private void AppendRequirementDuplicateMergeRanges(
            MergeCells mergeCells,
            IReadOnlyList<Dictionary<string, object>> rows,
            IReadOnlyList<string> columnOrder
        )
        {
            if (mergeCells == null || rows == null || rows.Count < 2 || columnOrder == null || columnOrder.Count == 0)
            {
                return;
            }

            int l2IdColumnIndex = FindColumnIndex(columnOrder, "L2 REQ ID");
            if (l2IdColumnIndex < 0)
            {
                return;
            }

            var mergeColumnIndexes = RequirementMergeCandidateColumns
                .Select(columnName => new { Name = columnName, Index = FindColumnIndex(columnOrder, columnName) })
                .Where(item => item.Index >= 0)
                .ToList();
            if (mergeColumnIndexes.Count == 0)
            {
                return;
            }

            int groupStart = 0;
            string currentL2Id = GetComparableCellValue(rows[0], columnOrder[l2IdColumnIndex]);

            for (int i = 1; i <= rows.Count; i++)
            {
                string candidateL2Id = i < rows.Count
                    ? GetComparableCellValue(rows[i], columnOrder[l2IdColumnIndex])
                    : null;
                bool boundary = i == rows.Count || !string.Equals(currentL2Id, candidateL2Id, StringComparison.OrdinalIgnoreCase);
                if (!boundary)
                {
                    continue;
                }

                int groupEnd = i - 1;
                if (groupEnd > groupStart && !string.IsNullOrWhiteSpace(currentL2Id))
                {
                    foreach (var mergeColumn in mergeColumnIndexes)
                    {
                        AppendColumnMergeRangesInGroup(
                            mergeCells,
                            rows,
                            columnOrder[mergeColumn.Index],
                            mergeColumn.Index,
                            groupStart,
                            groupEnd
                        );
                    }
                }

                groupStart = i;
                currentL2Id = candidateL2Id;
            }
        }

        private void AppendColumnMergeRangesInGroup(
            MergeCells mergeCells,
            IReadOnlyList<Dictionary<string, object>> rows,
            string columnName,
            int columnIndex,
            int groupStart,
            int groupEnd
        )
        {
            int runStart = groupStart;
            string runValue = GetComparableCellValue(rows[groupStart], columnName);

            for (int rowOffset = groupStart + 1; rowOffset <= groupEnd + 1; rowOffset++)
            {
                string candidateValue = rowOffset <= groupEnd
                    ? GetComparableCellValue(rows[rowOffset], columnName)
                    : null;
                bool boundary = rowOffset > groupEnd || !string.Equals(runValue, candidateValue, StringComparison.OrdinalIgnoreCase);
                if (!boundary)
                {
                    continue;
                }

                int runEnd = rowOffset - 1;
                if (runEnd > runStart && !string.IsNullOrWhiteSpace(runValue))
                {
                    int startRowNumber = runStart + 2; // Header is row 1.
                    int endRowNumber = runEnd + 2;
                    string columnLetter = _spreadsheetService.GetColumnLetter(columnIndex + 1);
                    mergeCells.Append(new MergeCell
                    {
                        Reference = new StringValue(
                            $"{columnLetter}{startRowNumber}:{columnLetter}{endRowNumber}"
                        )
                    });
                }

                runStart = rowOffset;
                runValue = candidateValue;
            }
        }

        private static int FindColumnIndex(IReadOnlyList<string> columnOrder, string columnName)
        {
            for (int i = 0; i < columnOrder.Count; i++)
            {
                if (string.Equals(columnOrder[i], columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static string GetComparableCellValue(
            IReadOnlyDictionary<string, object> row,
            string columnName
        )
        {
            if (row == null || string.IsNullOrWhiteSpace(columnName))
            {
                return string.Empty;
            }

            foreach (var kvp in row)
            {
                if (!string.Equals(kvp.Key, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return Convert.ToString(kvp.Value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
            }

            return string.Empty;
        }
    }
}
