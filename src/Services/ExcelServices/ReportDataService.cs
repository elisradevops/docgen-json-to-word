using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace JsonToWord.Services.ExcelServices
{
    public class ReportDataService : IReportDataService
    {
        private readonly ILogger<ReportDataService> _logger;
        private readonly IStylesheetService _stylesheetService;
        private readonly ISpreadsheetService _spreadsheetService;
        private readonly IExcelHelperService _excelHelperService;
        private WorksheetPart _currentWorksheetPart;

        public ReportDataService(ILogger<ReportDataService> logger, IStylesheetService stylesheetService, ISpreadsheetService spreadsheetService, IExcelHelperService excelHelperService) { 
            _logger = logger;
            _stylesheetService = stylesheetService;
            _spreadsheetService = spreadsheetService;
            _excelHelperService = excelHelperService;
        }

        public void AddDataRows(SheetData sheetData, MergeCells mergeCells,
                List<TestSuiteModel> testSuites,
                List<ColumnDefinition> columnDefinitions,
                Dictionary<string, int> columnCountForeachGroup,
                ref uint rowIndex, WorksheetPart worksheetPart, bool groupBySuite)
        {
            if (sheetData == null)
                throw new ArgumentNullException(nameof(sheetData));
            if (mergeCells == null)
                throw new ArgumentNullException(nameof(mergeCells));
            if (testSuites == null)
                throw new ArgumentNullException(nameof(testSuites));
            if (columnDefinitions == null || !columnDefinitions.Any())
                throw new ArgumentException("Column definitions cannot be null or empty", nameof(columnDefinitions));
            if (worksheetPart == null)
                throw new ArgumentNullException(nameof(worksheetPart));

            _currentWorksheetPart = worksheetPart;
            try
            {
                _logger.LogInformation("Starting to add data rows. Group by suite: {GroupBySuite}, Initial row index: {RowIndex}",
               groupBySuite, rowIndex);
                bool useAlternateColor = false;
                uint suiteTitleStyleIndex = 2;
                uint dataStyleIndex1 = 6;
                uint dataStyleIndex2 = 7;
                uint dateStyleIndex1 = 8;
                uint dateStyleIndex2 = 9;
                uint numberStyleIndex1 = 10;
                uint numberStyleIndex2 = 11;
                uint columnIndexOffset = 0;
                foreach (var testSuite in testSuites)
                {
                    if (groupBySuite)  // Only add suite title row if groupBySuite is true
                    {
                        // Add suite title row
                        Row suiteRow = new Row { RowIndex = rowIndex++ };
                        sheetData.Append(suiteRow);

                        // Create suite title cell
                        Cell suiteCell = new Cell
                        {
                            CellReference = $"{_spreadsheetService.GetColumnLetter(1)}{suiteRow.RowIndex}",
                            CellValue = new CellValue($"Suite: {testSuite.SuiteName}"),
                            DataType = CellValues.String,
                            StyleIndex = suiteTitleStyleIndex
                        };
                        suiteRow.Append(suiteCell);

                        // Merge cells for suite title across all columns
                        mergeCells.Append(new MergeCell
                        {
                            Reference = new StringValue(
                                $"{_spreadsheetService.GetColumnLetter(1)}{suiteRow.RowIndex}:{_spreadsheetService.GetColumnLetter(columnDefinitions.Count)}{suiteRow.RowIndex}"
                            )
                        });
                    }


                    var testCaseColumnDefinitions = columnDefinitions.Where(x => x.Group == "Test Cases").ToList();
                    var requirementsColumnDefinitions = columnDefinitions.Where(x => x.Group == "Requirements").ToList();
                    var bugsColumnDefinitions = columnDefinitions.Where(x => x.Group == "Bugs").ToList();
                    var crsColumnDefinitions = columnDefinitions.Where(x => x.Group == "CRs").ToList();

                    // Calculate fixed column start positions for each group
                    int testCaseColStart = 0;
                    int reqColStart = testCaseColumnDefinitions.Count;
                    int bugColStart = reqColStart + requirementsColumnDefinitions.Count;
                    int crColStart = bugColStart + bugsColumnDefinitions.Count;

                    // Add test cases
                    foreach (var testCase in testSuite.TestCases)
                    {
                        // Alternate background color for each test case
                        useAlternateColor = !useAlternateColor;
                        uint currentDataStyleIndex = useAlternateColor ? dataStyleIndex1 : dataStyleIndex2;
                        uint currentDateStyleIndex = useAlternateColor ? dateStyleIndex1 : dateStyleIndex2;
                        uint currentNumberStyleIndex = useAlternateColor ? numberStyleIndex1 : numberStyleIndex2;

                        // We'll use this to track the current row as we add data
                        uint currentRowIndex = rowIndex;
                        
                        // Determine the number of rows needed for the test case part (steps or single row)
                        int testCaseStepRows = testCase.TestSteps?.Count ?? 0;
                        int testCaseHistoryRows = testCase.HistoryEntries?.Count ?? 0;
                        int testCaseRows = Math.Max(testCaseStepRows, testCaseHistoryRows);

                        // Determine the maximum number of rows needed for any associated item group
                        int maxAssociatedItems = Math.Max(testCase.AssociatedRequirements?.Count ?? 0,
                            Math.Max(testCase.AssociatedBugs?.Count ?? 0, testCase.AssociatedCRs?.Count ?? 0));

                        // The total row span for this test case block is the max of the two
                        int rowSpan = Math.Max(testCaseRows, maxAssociatedItems);
                        if (rowSpan == 0) rowSpan = 1; // Ensure at least one row is always processed

                        // Emit rows (step rows + history rows). History entries align with row index (start at first data row).
                        bool isFirstRow = true;
                        for (int rowOffset = 0; rowOffset < rowSpan; rowOffset++)
                        {
                            TestStepModel step =
                                testCase.TestSteps != null && rowOffset < testCase.TestSteps.Count
                                    ? testCase.TestSteps[rowOffset]
                                    : null;
                            string historyEntry =
                                testCase.HistoryEntries != null && rowOffset < testCase.HistoryEntries.Count
                                    ? testCase.HistoryEntries[rowOffset]
                                    : null;

                            try
                            {
                                Row row = GetOrCreateRow(sheetData, currentRowIndex++);
                                row.OutlineLevel = (ByteValue)(groupBySuite ? 1 : 0);  // Only outline if grouping by suite

                                // Create cells for test case columns
                                for (int i = 0; i < testCaseColumnDefinitions.Count; i++)
                                {
                                    string colLetter = _spreadsheetService.GetColumnLetter(testCaseColStart + i + 1);
                                    string cellRef = $"{colLetter}{row.RowIndex}";
                                    string property = testCaseColumnDefinitions[i].Property;

                                    try
                                    {
                                        if (property == "SuiteName" && !groupBySuite)
                                        {
                                            SafeAddCell(row, cellRef, testSuite.SuiteName, currentDataStyleIndex);
                                        }
                                        else if (isFirstRow)
                                        {
                                            AddTestCaseCellWithSteps(
                                                row,
                                                property,
                                                cellRef,
                                                testCase,
                                                step,
                                                historyEntry,
                                                currentDataStyleIndex,
                                                currentDateStyleIndex,
                                                currentNumberStyleIndex
                                            );
                                        }
                                        else
                                        {
                                            AddStepOnlyCell(row, property, cellRef, step, historyEntry, currentDataStyleIndex);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        _logger.LogError(ex, "Error processing cell {CellRef} for property {Property}",
                                            cellRef, property);
                                        // Add an error indicator to the cell
                                        SafeAddCell(row, cellRef, "Error", currentDataStyleIndex);
                                    }
                                }
                                isFirstRow = false;
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, "Error processing test row for test case {TestCaseId}",
                                    testCase.TestCaseId);
                            }
                        }

                        // Fill empty cells for remaining rows in the span for the test case columns
                        uint rowsHandledForTestCase = currentRowIndex - rowIndex;
                        for (uint i = rowsHandledForTestCase; i < rowSpan; i++)
                        {
                            Row row = GetOrCreateRow(sheetData, rowIndex + i);
                            for (int j = 0; j < testCaseColumnDefinitions.Count; j++)
                            {
                                string colLetter = _spreadsheetService.GetColumnLetter(testCaseColStart + j + 1);
                                string cellRef = $"{colLetter}{row.RowIndex}";
                                SafeAddCell(row, cellRef, string.Empty, currentDataStyleIndex);
                            }
                        }

                        // Process requirements, bugs, and CRs, starting from the top of the block
                        ProcessAssociatedItems(sheetData, testCase.AssociatedRequirements, requirementsColumnDefinitions, reqColStart, rowIndex, rowSpan, currentDataStyleIndex, currentNumberStyleIndex, currentDateStyleIndex, "Requirement");
                        ProcessAssociatedItems(sheetData, testCase.AssociatedBugs, bugsColumnDefinitions, bugColStart, rowIndex, rowSpan, currentDataStyleIndex, currentNumberStyleIndex, currentDateStyleIndex, "Bug");
                        ProcessAssociatedItems(sheetData, testCase.AssociatedCRs, crsColumnDefinitions, crColStart, rowIndex, rowSpan, currentDataStyleIndex, currentNumberStyleIndex, currentDateStyleIndex, "CR");

                        // Merge the test case cells
                        if (rowSpan > 1)
                        {
                            var stepProperties = new HashSet<string> { "StepNo", "StepAction", "StepExpected", "StepRunStatus", "StepErrorMessage", "History" };

                            // Merge cells for test case columns that are not step-specific
                            for (int i = 0; i < testCaseColumnDefinitions.Count; i++)
                            {
                                string property = testCaseColumnDefinitions[i].Property;
                                if (!stepProperties.Contains(property))
                                {
                                    string colLetter = _spreadsheetService.GetColumnLetter(testCaseColStart + i + 1);
                                    string startCellRef = $"{colLetter}{rowIndex}";
                                    string endCellRef = $"{colLetter}{rowIndex + (uint)rowSpan - 1}";
                                    string mergeRef = $"{startCellRef}:{endCellRef}";
                                    mergeCells.Append(new MergeCell
                                    {
                                        Reference = new StringValue(mergeRef)
                                    });
                                }
                            }
                        }

                        // Advance the main row index for the next test case
                        rowIndex += (uint)rowSpan;
                    }
                }

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while adding data rows");
                throw; // Re-throw to allow caller to handle
            }
        }

        private void ProcessAssociatedItems<T>(SheetData sheetData, List<T> items, List<ColumnDefinition> columnDefinitions, int colStart, uint startRowIndex, int rowSpan, uint dataStyleIndex, uint numberStyleIndex, uint dateStyleIndex, string itemType) where T : AssociatedItemModel
        {
            if (columnDefinitions.Count == 0) return;

            for (int i = 0; i < rowSpan; i++)
            {
                uint currentRow = startRowIndex + (uint)i;
                Row row = GetOrCreateRow(sheetData, currentRow);
                var item = (items != null && i < items.Count) ? items[i] : null;

                for (int j = 0; j < columnDefinitions.Count; j++)
                {
                    string colLetter = _spreadsheetService.GetColumnLetter(colStart + j + 1);
                    string cellRef = $"{colLetter}{row.RowIndex}";

                    if (item != null)
                    {
                        string property = columnDefinitions[j].Property;
                        AddAssociatedItemCell(row, property, cellRef, item, dataStyleIndex, numberStyleIndex, dateStyleIndex, itemType);
                    }
                    else
                    {
                        FillEmptyCell(row, cellRef, dataStyleIndex);
                    }
                }
            }
        }

        private Row GetOrCreateRow(SheetData sheetData, uint rowIndex)
        {
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            return row;
        }

        private void AddTestCaseCell(Row row, string property, string cellRef, TestCaseModel testCase, uint dataStyleIndex, uint dateStyleIndex, uint numberStyleIndex)
        {
            // Handle step-independent properties only
            AddTestCasePropertyCell(row, property, cellRef, testCase, null,
                                   dataStyleIndex, dateStyleIndex, numberStyleIndex);
        }

        private void SafeAddCell(Row row, string cellRef, string value, uint styleIndex)
        {
            SafeAddCell(row, cellRef, value, styleIndex, CellValues.String);
        }

        private void SafeAddCell(Row row, string cellRef, string value, uint styleIndex, CellValues dataType)
        {
            try
            {
                var cell = new Cell
                {
                    CellReference = cellRef,
                    CellValue = new CellValue(value ?? string.Empty),
                    DataType = dataType,
                    StyleIndex = styleIndex
                };
                row.Append(cell);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error adding cell at reference {CellRef}", cellRef);
                throw;
            }
        }

        private void AddTestCaseCellWithSteps(Row row, string property, string cellRef,
                                TestCaseModel testCase, TestStepModel step, string historyEntry,
                                uint dataStyleIndex, uint dateStyleIndex, uint numberStyleIndex)
        {
            // If it's a step-specific property, handle it directly
            if (property == "StepNo")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, step?.StepNo ?? string.Empty, dataStyleIndex));
            else if (property == "StepAction")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, StripHtmlAndTruncate(step?.StepAction), dataStyleIndex));
            else if (property == "StepExpected")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, StripHtmlAndTruncate(step?.StepExpected), dataStyleIndex));
            else if (property == "StepRunStatus")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, step?.StepRunStatus ?? string.Empty, dataStyleIndex));
            else if (property == "StepErrorMessage")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, step?.StepErrorMessage ?? string.Empty, dataStyleIndex));
            else if (property == "History")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, StripHtmlAndTruncate(historyEntry), dataStyleIndex));
            else
                // For all test case properties, use the common method
                AddTestCasePropertyCell(row, property, cellRef, testCase, step,
                                       dataStyleIndex, dateStyleIndex, numberStyleIndex);
        }

        private void AddAssociatedItemCell(Row row, string property, string cellRef,
                                           AssociatedItemModel associatedItem, uint dataStyleIndex, uint numberStyleIndex, uint dateStyleIndex, string type)
        {
             if (property == $"{type}Id")
            {
                if (!string.IsNullOrEmpty(associatedItem.Url))
                    row.Append(_spreadsheetService.CreateHyperlinkCell(_currentWorksheetPart, cellRef, associatedItem.Id.ToString(), associatedItem.Url,
                             numberStyleIndex, $"Open {type} {associatedItem.Id} in Azure Devops"));
                else
                    row.Append(_spreadsheetService.CreateNumberCell(cellRef, associatedItem.Id.ToString(), numberStyleIndex));
            }
            else if (property == $"{type}Name")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, associatedItem.Title, dataStyleIndex));
            else
            {
                // Check if this is a custom field (dynamic property)
                // Convert to camelCase for dictionary lookup
                string fieldName = char.ToLower(property[0]) + property.Substring(1);
                string value = GetCustomFieldValue(associatedItem.CustomFields, fieldName);

                if (!string.IsNullOrEmpty(value))
                {
                    // Check if the field might be a date
                    if (fieldName.Contains("date") && DateTime.TryParse(value, out DateTime dateValue))
                    {
                        row.Append(_spreadsheetService.CreateDateCell(cellRef, dateValue, dateStyleIndex));
                    }
                    else
                    {
                        row.Append(_spreadsheetService.CreateTextCell(cellRef, value, dataStyleIndex));
                    }
                }
                else
                {
                    row.Append(_spreadsheetService.CreateTextCell(cellRef, "", dataStyleIndex));
                }
            }
        }

        private void AddTestCasePropertyCell(Row row, string property, string cellRef,
                                           TestCaseModel testCase, TestStepModel step,
                                           uint dataStyleIndex, uint dateStyleIndex, uint numberStyleIndex)
        {
            if (property == "TestCaseId")
            {
                if (!string.IsNullOrEmpty(testCase.TestCaseUrl))
                    row.Append(_spreadsheetService.CreateHyperlinkCell(_currentWorksheetPart, cellRef, testCase.TestCaseId.ToString(), testCase.TestCaseUrl,
                             numberStyleIndex, $"Open test case {testCase.TestCaseId} in Azure Devops"));
                else
                    row.Append(_spreadsheetService.CreateNumberCell(cellRef, testCase.TestCaseId.ToString(), numberStyleIndex));
            }
            else if (property == "TestCaseName")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.TestCaseName, dataStyleIndex));
            else if (property == "ExecutionDate" && !string.IsNullOrEmpty(testCase.ExecutionDate))
                row.Append(_spreadsheetService.CreateDateCell(cellRef, DateTime.Parse(testCase.ExecutionDate), dateStyleIndex));
            else if (property == "TestCaseResult")
            {
                if (testCase.TestCaseResult != null && !string.IsNullOrEmpty(testCase.TestCaseResult.Url))
                    row.Append(_spreadsheetService.CreateHyperlinkCell(_currentWorksheetPart, cellRef, testCase.TestCaseResult.ResultMessage,
                             testCase.TestCaseResult.Url, dataStyleIndex,
                             $"Open last run result for test case {testCase.TestCaseId} in Azure Devops"));
                else if (testCase.TestCaseResult != null)
                    row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.TestCaseResult.ResultMessage, dataStyleIndex));
                else
                    row.Append(_spreadsheetService.CreateTextCell(cellRef, "", dataStyleIndex));
            }
            else if (property == "FailureType")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.FailureType, dataStyleIndex));
            else if (property == "TestCaseComment")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.Comment, dataStyleIndex));
            else if (property == "RunBy")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.RunBy, dataStyleIndex));
            else if (property == "Configuration")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.Configuration, dataStyleIndex));
            else if (property == "State")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.State, dataStyleIndex));
            else if (property == "StateChangeDate")
            {
                if (!string.IsNullOrEmpty(testCase.StateChangeDate) && DateTime.TryParse(testCase.StateChangeDate, out DateTime scd))
                    row.Append(_spreadsheetService.CreateDateCell(cellRef, scd, dateStyleIndex));
                else
                    row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.StateChangeDate, dataStyleIndex));
            }
            else if (property == "AssociatedRequirementCount")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.AssociatedRequirements?.Count.ToString(), dateStyleIndex));
            else if (property.StartsWith("AssociatedRequirement_"))
                HandleRequirementCell(row, property, cellRef, testCase, dataStyleIndex);
            else if (property == "AssociatedBugCount")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.AssociatedBugs?.Count.ToString(), dateStyleIndex));
            else if (property.StartsWith("AssociatedBug_"))
                HandleBugCell(row, property, cellRef, testCase, dataStyleIndex);
            else if (property == "AssociatedCRCount")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, testCase.AssociatedCRs?.Count.ToString(), dateStyleIndex));
            else if (property.StartsWith("AssociatedCR_"))
                HandleCRCell(row, property, cellRef, testCase, dataStyleIndex);
            else
            {
                // Check if this is a custom field (dynamic property)
                // Convert to camelCase for dictionary lookup
                string fieldName = char.ToLower(property[0]) + property.Substring(1);
                string value = GetCustomFieldValue(testCase.CustomFields, fieldName);

                if (!string.IsNullOrEmpty(value))
                {
                    // Check if the field might be a date
                    if (fieldName.Contains("date") && DateTime.TryParse(value, out DateTime dateValue))
                    {
                        row.Append(_spreadsheetService.CreateDateCell(cellRef, dateValue, dateStyleIndex));
                    }
                    else
                    {
                        row.Append(_spreadsheetService.CreateTextCell(cellRef, value, dataStyleIndex));
                    }
                }
                else
                {
                    FillEmptyCell(row, cellRef, dataStyleIndex);
                }
            }
        }

        private void FillEmptyCell(Row row, string cellRef, uint dataStyleIndex)
        {
            row.Append(_spreadsheetService.CreateTextCell(cellRef, "", dataStyleIndex));
        }

        
        private void HandleRequirementCell(Row row, string property, string cellRef, TestCaseModel testCase, uint dataStyleIndex)
        {
            if (int.TryParse(property.Substring("AssociatedRequirement_".Length), out int reqIndex))
            {
                if (testCase.AssociatedRequirements != null &&
                    reqIndex < testCase.AssociatedRequirements.Count &&
                    testCase.AssociatedRequirements[reqIndex] != null)
                {
                    var req = testCase.AssociatedRequirements[reqIndex];

                    if (!string.IsNullOrEmpty(req.Url))
                        row.Append(_spreadsheetService.CreateHyperlinkCell(_currentWorksheetPart, cellRef, $"{req.Id} {req.Title}",
                                req.Url, dataStyleIndex, $"Open Requirement {req.Id} in Azure DevOps"));
                    else
                        row.Append(_spreadsheetService.CreateTextCell(cellRef, $"{req.Id} {req.Title}", dataStyleIndex));
                }
                else
                {
                    FillEmptyCell(row, cellRef, dataStyleIndex);
                }
            }
            else
            {
                FillEmptyCell(row, cellRef, dataStyleIndex);
            }
        }

        private void HandleBugCell(Row row, string property, string cellRef, TestCaseModel testCase, uint dataStyleIndex)
        {
            if (int.TryParse(property.Substring("AssociatedBug_".Length), out int bugIdx))
            {
                if (testCase.AssociatedBugs != null &&
                    bugIdx < testCase.AssociatedBugs.Count &&
                    testCase.AssociatedBugs[bugIdx] != null)
                {
                    var bug = testCase.AssociatedBugs[bugIdx];

                    if (!string.IsNullOrEmpty(bug.Url))
                        row.Append(_spreadsheetService.CreateHyperlinkCell(_currentWorksheetPart, cellRef, $"{bug.Id} {bug.Title}",
                                bug.Url, dataStyleIndex, $"Open Bug {bug.Id} in Azure DevOps"));
                    else
                        row.Append(_spreadsheetService.CreateTextCell(cellRef, $"{bug.Id} {bug.Title}", dataStyleIndex));
                }
                else
                {
                    FillEmptyCell(row, cellRef, dataStyleIndex);
                }
            }
            else
            {
                FillEmptyCell(row, cellRef, dataStyleIndex);
            }
        }

        private void HandleCRCell(Row row, string property, string cellRef, TestCaseModel testCase, uint dataStyleIndex)
        {
            if (int.TryParse(property.Substring("AssociatedCR_".Length), out int crIdx))
            {
                if (testCase.AssociatedCRs != null &&
                    crIdx < testCase.AssociatedCRs.Count &&
                    testCase.AssociatedCRs[crIdx] != null)
                {
                    var cr = testCase.AssociatedCRs[crIdx];
                    if (!string.IsNullOrEmpty(cr.Url))
                        row.Append(_spreadsheetService.CreateHyperlinkCell(_currentWorksheetPart, cellRef, $"{cr.Id} {cr.Title}",
                                cr.Url, dataStyleIndex, $"Open Change Request {cr.Id} in Azure DevOps"));
                    else
                        row.Append(_spreadsheetService.CreateTextCell(cellRef, $"{cr.Id} {cr.Title}", dataStyleIndex));
                }
                else
                {
                    FillEmptyCell(row, cellRef, dataStyleIndex);
                }
            }
            else
            {
                FillEmptyCell(row, cellRef, dataStyleIndex);
            }
        }

        private string GetCustomFieldValue(Dictionary<string, object> customFields, string fieldName)
        {
            if (customFields != null && customFields.TryGetValue(fieldName, out var fieldValue) && fieldValue != null)
            {
                return _excelHelperService.GetValueString(fieldValue);
            }
            return string.Empty;
        }


        private void AddStepOnlyCell(Row row, string property, string cellRef, TestStepModel step, string historyEntry, uint dataStyleIndex)
        {
            if (property == "StepNo")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, step?.StepNo ?? string.Empty, dataStyleIndex));
            else if (property == "StepAction")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, StripHtmlAndTruncate(step?.StepAction), dataStyleIndex));
            else if (property == "StepExpected")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, StripHtmlAndTruncate(step?.StepExpected), dataStyleIndex));
            else if (property == "StepRunStatus")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, step?.StepRunStatus ?? string.Empty, dataStyleIndex));
            else if (property == "StepErrorMessage")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, step?.StepErrorMessage ?? string.Empty, dataStyleIndex));
            else if (property == "History")
                row.Append(_spreadsheetService.CreateTextCell(cellRef, StripHtmlAndTruncate(historyEntry), dataStyleIndex));
            else
                FillEmptyCell(row, cellRef, dataStyleIndex);
        }

        private string StripHtmlAndTruncate(string html, int maxLength = 500)
        {
            if (string.IsNullOrEmpty(html))
                return string.Empty;

            try
            {
                // Remove HTML tags using regex
                string plainText = Regex.Replace(html, "<[^>]*>", string.Empty);

                // Replace common HTML entities
                plainText = plainText.Replace("&nbsp;", " ")
                                    .Replace("&lt;", "<")
                                    .Replace("&gt;", ">")
                                    .Replace("&amp;", "&")
                                    .Replace("&quot;", "\"");

                // Normalize whitespace
                plainText = Regex.Replace(plainText, @"\s+", " ").Trim();

                // Truncate if too long
                if (plainText.Length > maxLength)
                    return plainText.Substring(0, maxLength - 3) + "...";

                return plainText;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error stripping HTML from content");
                return "Error parsing HTML content";
            }
        }

    }
}
