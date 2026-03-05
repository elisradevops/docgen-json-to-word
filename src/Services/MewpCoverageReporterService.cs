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
            "SR #",
            "L2 REQ Title",
            "L2 Owner",
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
            "SR #",
            "L2 REQ Title",
            "L2 Owner",
            "L2 SubSystem",
            "L2 Run Status",
            "L3 REQ ID",
            "L3 REQ Title",
        };

        private static readonly string[] L2CoverageSummaryColumns = new[]
        {
            "SR num",
            "L2 REQ Title",
            "L2 Run Status",
            "L2 Owner",
        };

        private static readonly HashSet<string> BugColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Bug ID",
            "Bug Title",
            "Bug Responsibility",
        };

        private static readonly HashSet<string> L3Columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "L3 REQ ID",
            "L3 REQ Title",
        };

        private static readonly HashSet<string> L4Columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
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

                _stylesheetService.EnsureStylesheet(workbookPart);

                var columnOrder =
                    coverageModel.ColumnOrder != null && coverageModel.ColumnOrder.Count > 0
                        ? coverageModel.ColumnOrder.ToArray()
                        : DefaultColumnOrder;
                var rows = coverageModel.Rows ?? new List<Dictionary<string, object>>();

                InsertCoverageWorksheet(
                    workbookPart,
                    safeName,
                    columnOrder,
                    rows,
                    coverageModel.MergeDuplicateRequirementCells
                );
                InsertL2CoverageSummaryWorksheet(workbookPart, safeName, rows);
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

        private void InsertCoverageWorksheet(
            WorkbookPart workbookPart,
            string worksheetName,
            IReadOnlyList<string> columnOrder,
            IReadOnlyList<Dictionary<string, object>> rows,
            bool mergeDuplicateRequirementCells
        )
        {
            WorksheetPart worksheetPart = _spreadsheetService.GetOrCreateWorksheetPart(workbookPart, worksheetName);
            worksheetPart.Worksheet = new Worksheet();
            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet.Append(sheetData);

            SheetViews sheetViews = new SheetViews(
                new SheetView { WorkbookViewId = 0U, RightToLeft = false }
            );
            worksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);

            var columns = CreateColumns(columnOrder.ToArray());
            worksheetPart.Worksheet.InsertBefore(columns, sheetData);
            MergeCells mergeCells = new MergeCells();

            uint rowIndex = 1;
            Row headerRow = new Row { RowIndex = rowIndex };
            sheetData.Append(headerRow);
            for (int i = 0; i < columnOrder.Count; i++)
            {
                string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                Cell headerCell = _spreadsheetService.CreateTextCell(
                    $"{columnLetter}{rowIndex}",
                    columnOrder[i],
                    1
                );
                headerRow.Append(headerCell);
            }

            var useFirstAlternatingColorByRow = BuildRowAlternatingColorFlags(
                rows,
                columnOrder,
                mergeDuplicateRequirementCells
            );
            var duplicateFlagsByRow = BuildDuplicateHighlightFlags(rows);
            for (int rowOffset = 0; rowOffset < rows.Count; rowOffset++)
            {
                rowIndex++;
                var row = rows[rowOffset];
                Row dataRow = new Row { RowIndex = rowIndex };
                sheetData.Append(dataRow);
                bool useFirstAlternatingColor =
                    rowOffset >= 0 && rowOffset < useFirstAlternatingColorByRow.Count
                        ? useFirstAlternatingColorByRow[rowOffset]
                        : rowIndex % 2 == 0;
                var duplicateFlags =
                    rowOffset >= 0 && rowOffset < duplicateFlagsByRow.Count
                        ? duplicateFlagsByRow[rowOffset]
                        : new DuplicateHighlightFlags();

                for (int i = 0; i < columnOrder.Count; i++)
                {
                    string columnName = columnOrder[i];
                    string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                    string cellRef = $"{columnLetter}{rowIndex}";
                    object value = null;
                    string cellValue = string.Empty;
                    if (row != null && row.TryGetValue(columnName, out value))
                    {
                        cellValue = value?.ToString() ?? string.Empty;
                    }
                    uint styleIndex = ResolveDataStyle(columnName, useFirstAlternatingColor, duplicateFlags);
                    Cell cell;
                    if (IsNumericColumn(columnName) && TryToNumberCellValue(value, out string numericValue))
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

            if (mergeDuplicateRequirementCells)
            {
                AppendRequirementDuplicateMergeRanges(mergeCells, rows, columnOrder);
            }

            if (mergeCells.Any())
            {
                worksheetPart.Worksheet.InsertAfter(mergeCells, sheetData);
            }

            worksheetPart.Worksheet.Save();
        }

        private void InsertL2CoverageSummaryWorksheet(
            WorkbookPart workbookPart,
            string mainWorksheetName,
            IReadOnlyList<Dictionary<string, object>> rows
        )
        {
            string summaryWorksheetName = BuildSummaryWorksheetName(mainWorksheetName);
            var summaryRows = BuildL2CoverageSummaryRows(rows);
            WorksheetPart worksheetPart = _spreadsheetService.GetOrCreateWorksheetPart(workbookPart, summaryWorksheetName);
            worksheetPart.Worksheet = new Worksheet();
            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet.Append(sheetData);

            SheetViews sheetViews = new SheetViews(
                new SheetView { WorkbookViewId = 0U, RightToLeft = false }
            );
            worksheetPart.Worksheet.InsertBefore(sheetViews, sheetData);

            var columns = CreateColumns(L2CoverageSummaryColumns);
            worksheetPart.Worksheet.InsertBefore(columns, sheetData);

            uint rowIndex = 1;
            Row headerRow = new Row { RowIndex = rowIndex };
            sheetData.Append(headerRow);
            for (int i = 0; i < L2CoverageSummaryColumns.Length; i++)
            {
                string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                headerRow.Append(
                    _spreadsheetService.CreateTextCell(
                        $"{columnLetter}{rowIndex}",
                        L2CoverageSummaryColumns[i],
                        1
                    )
                );
            }

            for (int rowOffset = 0; rowOffset < summaryRows.Count; rowOffset++)
            {
                rowIndex++;
                bool useFirstAlternatingColor = rowOffset % 2 == 0;
                var sourceRow = summaryRows[rowOffset];
                Row dataRow = new Row { RowIndex = rowIndex };
                sheetData.Append(dataRow);

                for (int i = 0; i < L2CoverageSummaryColumns.Length; i++)
                {
                    string columnName = L2CoverageSummaryColumns[i];
                    string columnLetter = _spreadsheetService.GetColumnLetter(i + 1);
                    string cellRef = $"{columnLetter}{rowIndex}";
                    string cellValue = GetComparableCellValue(sourceRow, columnName);
                    uint styleIndex = ResolveDataStyle(columnName, useFirstAlternatingColor, new DuplicateHighlightFlags());
                    Cell cell = _spreadsheetService.CreateTextCell(cellRef, cellValue, styleIndex);
                    dataRow.Append(cell);
                }
            }

            worksheetPart.Worksheet.Save();
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
            if (key == "sr #") return 16;
            if (key == "sr num") return 16;
            if (key == "l2 req title") return 40;
            if (key == "l2 owner") return 20;
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

        private static uint ResolveDataStyle(
            string columnName,
            bool useFirstAlternatingColor,
            DuplicateHighlightFlags duplicateFlags
        )
        {
            bool isBugDuplicate = duplicateFlags?.BugDuplicate == true && IsBugColumn(columnName);
            bool isL3Duplicate = duplicateFlags?.L3Duplicate == true && IsL3Column(columnName);
            bool isL4Duplicate = duplicateFlags?.L4Duplicate == true && IsL4Column(columnName);

            if (IsNumericColumn(columnName))
            {
                if (isBugDuplicate)
                {
                    return useFirstAlternatingColor ? 28U : 29U;
                }
                if (IsBugColumn(columnName))
                {
                    return useFirstAlternatingColor ? 20U : 21U;
                }
                return useFirstAlternatingColor ? 10U : 11U;
            }

            if (isBugDuplicate)
            {
                return useFirstAlternatingColor ? 26U : 27U;
            }

            if (isL3Duplicate)
            {
                return useFirstAlternatingColor ? 30U : 31U;
            }

            if (isL4Duplicate)
            {
                return useFirstAlternatingColor ? 32U : 33U;
            }

            if (IsBugColumn(columnName))
            {
                return useFirstAlternatingColor ? 18U : 19U;
            }

            if (IsL3Column(columnName))
            {
                return useFirstAlternatingColor ? 22U : 23U;
            }

            if (IsL4Column(columnName))
            {
                return useFirstAlternatingColor ? 24U : 25U;
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

        private static bool IsL3Column(string columnName)
        {
            return !string.IsNullOrWhiteSpace(columnName) && L3Columns.Contains(columnName.Trim());
        }

        private static bool IsL4Column(string columnName)
        {
            return !string.IsNullOrWhiteSpace(columnName) && L4Columns.Contains(columnName.Trim());
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

        private string BuildSummaryWorksheetName(string mainWorksheetName)
        {
            const string summaryName = "MEWP L2 Summary";
            if (string.Equals(summaryName, mainWorksheetName?.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return "MEWP L2 Coverage Summary";
            }
            return summaryName;
        }

        private List<Dictionary<string, object>> BuildL2CoverageSummaryRows(
            IReadOnlyList<Dictionary<string, object>> sourceRows
        )
        {
            var summaryRows = new List<Dictionary<string, object>>();
            if (sourceRows == null || sourceRows.Count == 0)
            {
                return summaryRows;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in sourceRows)
            {
                var l2ReqId = GetComparableCellValue(row, "L2 REQ ID");
                var srNumber = GetComparableCellValue(row, "SR #");
                var uniqueKey = !string.IsNullOrWhiteSpace(l2ReqId) ? l2ReqId : srNumber;
                if (string.IsNullOrWhiteSpace(uniqueKey))
                {
                    continue;
                }

                if (!seen.Add(uniqueKey))
                {
                    continue;
                }
                var l2ReqTitle = GetComparableCellValue(row, "L2 REQ Full Title");
                if (string.IsNullOrWhiteSpace(l2ReqTitle))
                {
                    l2ReqTitle = GetComparableCellValue(row, "L2 REQ Title");
                }

                summaryRows.Add(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SR num"] = srNumber,
                    ["L2 REQ Title"] = l2ReqTitle,
                    ["L2 Run Status"] = GetComparableCellValue(row, "L2 Run Status"),
                    ["L2 Owner"] = GetComparableCellValue(row, "L2 Owner"),
                });
            }

            return summaryRows;
        }

        private List<DuplicateHighlightFlags> BuildDuplicateHighlightFlags(
            IReadOnlyList<Dictionary<string, object>> rows
        )
        {
            var result = new List<DuplicateHighlightFlags>();
            if (rows == null || rows.Count == 0)
            {
                return result;
            }

            var seenBugKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenL3Keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenL4Keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                var bugKey = ResolveBugDuplicateKey(row);
                var l3Key = ResolveL3DuplicateKey(row);
                var l4Key = ResolveL4DuplicateKey(row);

                result.Add(new DuplicateHighlightFlags
                {
                    BugDuplicate = !string.IsNullOrWhiteSpace(bugKey) && !seenBugKeys.Add(bugKey),
                    L3Duplicate = !string.IsNullOrWhiteSpace(l3Key) && !seenL3Keys.Add(l3Key),
                    L4Duplicate = !string.IsNullOrWhiteSpace(l4Key) && !seenL4Keys.Add(l4Key),
                });
            }

            return result;
        }

        private static string ResolveBugDuplicateKey(IReadOnlyDictionary<string, object> row)
        {
            var bugId = GetComparableCellValue(row, "Bug ID");
            if (!string.IsNullOrWhiteSpace(bugId))
            {
                return $"BUG:{bugId}";
            }

            var bugTitle = GetComparableCellValue(row, "Bug Title");
            var bugResponsibility = GetComparableCellValue(row, "Bug Responsibility");
            if (string.IsNullOrWhiteSpace(bugTitle) && string.IsNullOrWhiteSpace(bugResponsibility))
            {
                return string.Empty;
            }

            return $"BUG:{bugTitle}|{bugResponsibility}";
        }

        private static string ResolveL3DuplicateKey(IReadOnlyDictionary<string, object> row)
        {
            var l3ReqId = GetComparableCellValue(row, "L3 REQ ID");
            if (!string.IsNullOrWhiteSpace(l3ReqId))
            {
                return $"L3:{l3ReqId}";
            }

            var l3ReqTitle = GetComparableCellValue(row, "L3 REQ Title");
            if (string.IsNullOrWhiteSpace(l3ReqTitle))
            {
                return string.Empty;
            }

            return $"L3:{l3ReqTitle}";
        }

        private static string ResolveL4DuplicateKey(IReadOnlyDictionary<string, object> row)
        {
            var l4ReqId = GetComparableCellValue(row, "L4 REQ ID");
            if (!string.IsNullOrWhiteSpace(l4ReqId))
            {
                return $"L4:{l4ReqId}";
            }

            var l4ReqTitle = GetComparableCellValue(row, "L4 REQ Title");
            if (string.IsNullOrWhiteSpace(l4ReqTitle))
            {
                return string.Empty;
            }

            return $"L4:{l4ReqTitle}";
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

        private sealed class DuplicateHighlightFlags
        {
            public bool BugDuplicate { get; set; }
            public bool L3Duplicate { get; set; }
            public bool L4Duplicate { get; set; }
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
