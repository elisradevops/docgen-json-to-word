using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class TestReporterService : ITestReporterService
    {
        #region Fields
        private readonly ILogger<TestReporterService> _logger;
        private WorksheetPart _currentWorksheetPart;
        #endregion
        #region Constructor
        public TestReporterService(ILogger<TestReporterService> logger) {
            _logger = logger;
        }
        #endregion
        #region Interface Implementations
        public void Insert(SpreadsheetDocument document, string worksheetName, TestReporterModel testReporterModel)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name cannot be empty", nameof(worksheetName));

            if (testReporterModel == null)
                throw new ArgumentNullException(nameof(testReporterModel));

            try
            {
                // Get or add workbook part
                WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
                if (workbookPart.Workbook == null)
                    workbookPart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart if it doesn't exist
                WorksheetPart worksheetPart = GetOrCreateWorksheetPart(workbookPart, worksheetName);
                _currentWorksheetPart = worksheetPart; // Store the current worksheet part (moved after creation)
                if (worksheetPart.Worksheet == null)
                    worksheetPart.Worksheet = new Worksheet();

                // Set the spreadsheet view options
                SheetViews sheetViews = new SheetViews(
                    new SheetView { WorkbookViewId = 0U, RightToLeft = false }
                );
                worksheetPart.Worksheet.Append(sheetViews);

                // Define column definitions and widths
                var columnDefinitions = DefineColumns(testReporterModel);
                Columns columns = CreateColumns(columnDefinitions);
                worksheetPart.Worksheet.Append(columns);

                // Create SheetData for content
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet.Append(sheetData);

                // Prepare for merging cells
                MergeCells mergeCells = new MergeCells();
                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());

                // Add Stylesheet if it doesn't exist already
                EnsureStylesheet(workbookPart);

                // Create header row with column names
                CreateHeaderRow(sheetData, columnDefinitions);

                // Add data rows
                uint rowIndex = 2; // Start after header row
                AddDataRows(sheetData, mergeCells, testReporterModel.TestSuites, columnDefinitions, ref rowIndex, worksheetPart);

                // Assign the worksheet to the worksheetPart and save
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error inserting grouped table into worksheet '{worksheetName}': {ex.Message}");
                throw;
            }
        }
        #endregion
        #region Private Methods
        private WorksheetPart GetOrCreateWorksheetPart(WorkbookPart workbookPart, string worksheetName)
        {
            // Check if the worksheet already exists
            Sheets? sheets = workbookPart.Workbook.Sheets;
            if (sheets == null)
            {
                sheets = new Sheets();
                workbookPart.Workbook.Append(sheets);
            }

            // Look for existing worksheet
            Sheet? existingSheet = sheets.Elements<Sheet>()
                .FirstOrDefault(s => s.Name == worksheetName);

            if (existingSheet != null)
            {
                // Return the existing worksheet part - Renamed variable to avoid conflict
                WorksheetPart existingWorksheetPart = (WorksheetPart)workbookPart.GetPartById(existingSheet.Id!);
                // Make sure Worksheet is initialized
                if (existingWorksheetPart.Worksheet == null)
                    existingWorksheetPart.Worksheet = new Worksheet();
                return existingWorksheetPart;
            }

            // Create new worksheet part and register it
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            // Initialize the Worksheet
            worksheetPart.Worksheet = new Worksheet();

            Sheet sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = GetUniqueSheetId(sheets),
                Name = worksheetName
            };
            sheets.Append(sheet);

            return worksheetPart;
        }

        private uint GetUniqueSheetId(Sheets sheets)
        {
            uint maxSheetId = 0;
            foreach (Sheet sheet in sheets.Elements<Sheet>())
            {
                if (sheet.SheetId?.Value > maxSheetId)
                    maxSheetId = sheet.SheetId.Value;
            }
            return maxSheetId + 1;
        }

        private List<(string Name, int Width, string Property)> DefineColumns(TestReporterModel testReporterModel)
        {

            int maxRequirementCount = GetMaxRequirementCount(testReporterModel);
            int maxBugCount = GetMaxBugCount(testReporterModel);
            int maxCRCount = GetMaxCRCount(testReporterModel);

            var allColumns = new List<(string Name, int Width, string Property)>
            {
                ("Test Case ID", 15, "TestCaseId"),
                ("Test Case Title", 30, "TestCaseName"),

                // TestCase fields - include additional properties
                ("Execution Date", 16, "ExecutionDate"),
                ("TC Actual Result", 30, "TestCaseResult"),
                ("Failure Type", 15, "FailureType"),
                ("Test Case Comment", 30, "TestCaseComment"),

                // TestStep fields
                ("Step #", 10, "StepNo"),
                ("Step Action", 40, "StepAction"),
                ("Step Expected Result", 40, "StepExpected"),
                ("Step Actual Result", 30, "StepErrorMessage"),
                ("Step Run Status", 17, "StepRunStatus"),

                // TestCase fields - include additional properties
                ("Run By", 20, "RunBy"),
                ("Configuration", 15, "Configuration"),
                ("Automation Status", 18, "AutomationStatus"),
                ("Assigned To", 15, "AssignedTo"),
                ("Sub System", 15, "SubSystem"),
                ("Priority", 10, "Priority"),
                ("Associated Req. Count", 25, "AssociatedRequirementCount"),

            };

            // Add dynamic columns for each associated requirement
            for (int i = 0; i < maxRequirementCount; i++)
            {
                allColumns.Add(($"Associated Req. {i + 1}", 30, $"AssociatedRequirement_{i}"));
            }

            allColumns.Add(("Associated Bug Count", 30, "AssociatedBugCount"));

            for (int i = 0; i < maxBugCount; i++)
            {
                allColumns.Add(($"Associated Bug {i + 1}", 30, $"AssociatedBug_{i}"));
            }

            allColumns.Add(("Associated CR Count", 30, "AssociatedCRCount"));

            for (int i = 0; i < maxCRCount; i++)
            {
                allColumns.Add(($"Associated CR {i + 1}", 30, $"AssociatedCR_{i}"));
            }

            // Get list of columns that actually have data
            List<string> columnsWithData = GetColumnsWithData(testReporterModel);

            // Filter columns based on which ones have data
            return allColumns.Where(col => columnsWithData.Contains(col.Property)).ToList();
        }

        private int GetMaxRequirementCount(TestReporterModel testReporterModel)
        {
            int maxCount = 0;
            foreach (var suite in testReporterModel.TestSuites)
            {
                foreach (var testCase in suite.TestCases)
                {
                    if (testCase.AssociatedRequirements != null)
                    {
                        maxCount = Math.Max(maxCount, testCase.AssociatedRequirements.Count);
                    }
                }
            }
            return maxCount;
        }

        private int GetMaxBugCount(TestReporterModel testReporterModel)
        {
            int maxCount = 0;
            foreach (var suite in testReporterModel.TestSuites)
            {
                foreach (var testCase in suite.TestCases)
                {
                    if (testCase.AssociatedBugs != null)
                    {
                        maxCount = Math.Max(maxCount, testCase.AssociatedBugs.Count);
                    }
                }
            }
            return maxCount;
        }

        private int GetMaxCRCount(TestReporterModel testReporterModel)
        {
            int maxCount = 0;
            foreach (var suite in testReporterModel.TestSuites)
            {
                foreach (var testCase in suite.TestCases)
                {
                    if (testCase.AssociatedCRs != null)
                    {
                        maxCount = Math.Max(maxCount, testCase.AssociatedCRs.Count);
                    }

                }
            }
            return maxCount;
        }




        private Columns CreateColumns(List<(string Name, int Width, string Property)> columnDefinitions)
        {
            Columns columns = new Columns();
            uint columnIndex = 1;
            
            foreach (var col in columnDefinitions)
            {
                columns.Append(new Column { 
                    Min = columnIndex, 
                    Max = columnIndex++, 
                    Width = col.Width, 
                    CustomWidth = true 
                });
            }
            
            return columns;
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

        private List<string> GetColumnsWithData(TestReporterModel testReporterModel)
        {
            HashSet<string> columnsWithData = new HashSet<string>
            {
                "SuiteName",
                "TestCaseId",
                "TestCaseName"
            };

            // Set for tracking which requirement columns have data
            HashSet<string> reqColumnsWithData = new HashSet<string>();
            HashSet<string> bugColumnsWithData = new HashSet<string>();

            foreach (var suite in testReporterModel.TestSuites)
            {
                foreach (var testCase in suite.TestCases)
                {
                    // Existing columns check
                    if (!string.IsNullOrEmpty(testCase.ExecutionDate))
                        columnsWithData.Add("ExecutionDate");

                    if (testCase.TestCaseResult != null && !string.IsNullOrEmpty(testCase.TestCaseResult.ResultMessage))
                        columnsWithData.Add("TestCaseResult");


                    if (!string.IsNullOrEmpty(testCase.FailureType))
                        columnsWithData.Add("FailureType");

                    if (!string.IsNullOrEmpty(testCase.Comment))
                        columnsWithData.Add("TestCaseComment");

                    // Check steps data
                    if (testCase.TestSteps != null)
                    {
                        foreach (var step in testCase.TestSteps)
                        {
                            if (!string.IsNullOrEmpty(step.StepNo))
                                columnsWithData.Add("StepNo");
                            if (!string.IsNullOrEmpty(step.StepAction))
                                columnsWithData.Add("StepAction");
                            if (!string.IsNullOrEmpty(step.StepExpected))
                                columnsWithData.Add("StepExpected");
                            if (!string.IsNullOrEmpty(step.StepRunStatus))
                                columnsWithData.Add("StepRunStatus");
                            if (!string.IsNullOrEmpty(step.StepErrorMessage))
                                columnsWithData.Add("StepErrorMessage");
                        }
                    }

                    if (!string.IsNullOrEmpty(testCase.RunBy))
                        columnsWithData.Add("RunBy");
                    if (!string.IsNullOrEmpty(testCase.Configuration))
                        columnsWithData.Add("Configuration");
                    if (!string.IsNullOrEmpty(testCase.AutomationStatus))
                        columnsWithData.Add("AutomationStatus");
                    if(!string.IsNullOrEmpty(testCase.AssignedTo))
                        columnsWithData.Add("AssignedTo");
                    if(!string.IsNullOrEmpty(testCase.SubSystem))
                        columnsWithData.Add("SubSystem");
                    if (testCase.Priority.HasValue)
                        columnsWithData.Add("Priority");

                    // Check for each associated requirement and track which indexes have data
                    if (testCase.AssociatedRequirements != null)
                    {
                        if (testCase.AssociatedRequirements.Count > 0)
                        {
                            columnsWithData.Add("AssociatedRequirementCount");
                        }
                        for (int i = 0; i < testCase.AssociatedRequirements.Count; i++)
                        {
                            var req = testCase.AssociatedRequirements[i];
                            if (req != null && !string.IsNullOrEmpty(req.RequirementTitle))
                            {
                                reqColumnsWithData.Add($"AssociatedRequirement_{i}");
                            }
                        }
                    }

                    // Check for each associated bug and track which indexes have data
                    if (testCase.AssociatedBugs != null)
                    {
                        if (testCase.AssociatedBugs.Count > 0)
                        {
                            columnsWithData.Add("AssociatedBugCount");
                        }
                        for (int i = 0; i < testCase.AssociatedBugs.Count; i++)
                        {
                            var bug = testCase.AssociatedBugs[i];
                            if (bug != null && !string.IsNullOrEmpty(bug.BugTitle))
                            {
                                bugColumnsWithData.Add($"AssociatedBug_{i}");
                            }
                        }
                    }

                    // Check for each associated CR and track which indexes have data
                    if (testCase.AssociatedCRs != null)
                    {
                        if (testCase.AssociatedCRs.Count > 0)
                        {
                            columnsWithData.Add("AssociatedCRCount");
                        }
                        for (int i = 0; i < testCase.AssociatedCRs.Count; i++)
                        {
                            var cr = testCase.AssociatedCRs[i];
                            if (cr != null && !string.IsNullOrEmpty(cr.crTitle))
                            {
                                columnsWithData.Add($"AssociatedCR_{i}");
                            }
                        }
                    }
                }
            }

            // Add all requirement columns that have data
            columnsWithData.UnionWith(reqColumnsWithData);
            columnsWithData.UnionWith(bugColumnsWithData);

            return columnsWithData.ToList();
        }


        private void EnsureStylesheet(WorkbookPart workbookPart)
        {
            if (workbookPart.GetPartsOfType<WorkbookStylesPart>().Any())
                return;

            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateStylesheet();
            stylesPart.Stylesheet.Save();
        }

        private Stylesheet CreateStylesheet()
        {
            return new Stylesheet(
                new Fonts(
                new Font( // Index 0 - Default font
                        new FontSize { Val = 10 },
                        new FontName { Val = "Arial" }
                    ),
                    new Font( // Index 1 - Header font
                        new FontSize { Val = 11 },
                        new FontName { Val = "Arial" },
                        new Bold(),
                        new Color { Rgb = new HexBinaryValue("FFFFFFFF") }
                    ),
                    new Font( // Index 2 - SuiteName title font
                        new FontSize { Val = 11 },
                        new FontName { Val = "Arial" },
                        new Bold()
                    ),
                     new Font( // Index 3 - Hyperlink font
                        new FontSize { Val = 10 },
                        new FontName { Val = "Arial" },
                        new Underline { Val = UnderlineValues.Single },
                        new Color { Rgb = new HexBinaryValue("FF0563C1") } // Excel blue hyperlink color
                    )
                ),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }), // Index 0 - Default fill
                    new Fill(new PatternFill { PatternType = PatternValues.None }), // Index 1 - Not working
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FF000000") }) { PatternType = PatternValues.Solid }), // Index 2 - Black fill for headers
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FF0E2841") }) { PatternType = PatternValues.Solid }), // Index 3 - SuiteName title fill
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FFA6C9EC") }) { PatternType = PatternValues.Solid }), // Index 4 - First alternating color
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue("FFDAE9F8") }) { PatternType = PatternValues.Solid })  // Index 5 - Second alternating color
                ),
                new Borders(
                    new Border(), // Index 0 - Default border
                    new Border( // Index 1 - Thin border
                        new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin }
                    )
                ),
                new CellFormats(
                    new CellFormat(), // Index 0 - Default cell format
                    new CellFormat // Index 1 - Header format
                    {
                        FontId = 1,
                        FillId = 2,
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    },
                    new CellFormat // Index 2 - SuiteName title format
                    {
                        FontId = 1,
                        FillId = 3,
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Center }
                    },
                    new CellFormat(), // Index 3 - Reserved
                    new CellFormat(), // Index 4 - Reserved
                    new CellFormat(), // Index 5 - Reserved
                    new CellFormat // Index 6 - Data cell with first alternating color
                    {
                        BorderId = 1,
                        FillId = 4,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 7 - Data cell with second alternating color
                    {
                        BorderId = 1,
                        FillId = 5,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 8 - Date cell with first alternating color
                    {
                        BorderId = 1,
                        FillId = 4,
                        ApplyFill = true,
                        ApplyBorder = true,
                        NumberFormatId = 14, // Standard date format
                        ApplyNumberFormat = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 9 - Date cell with second alternating color
                    {
                        BorderId = 1,
                        FillId = 5,
                        ApplyFill = true,
                        ApplyBorder = true,
                        NumberFormatId = 14, // Standard date format
                        ApplyNumberFormat = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 10 - Number cell with first alternating color
                    {
                        BorderId = 1,
                        FillId = 4,
                        ApplyFill = true,
                        ApplyBorder = true,
                        NumberFormatId = 0, // General number format
                        ApplyNumberFormat = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 11 - Number cell with second alternating color
                    {
                        BorderId = 1,
                        FillId = 5,
                        ApplyFill = true,
                        ApplyBorder = true,
                        NumberFormatId = 0, // General number format
                        ApplyNumberFormat = true,
                        Alignment = new Alignment { Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 12 - Hyperlink style for first alternating color
                    {
                        FontId = 3, // Use the hyperlink font we defined
                        FillId = 4, // First alternating color fill
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    },
                    new CellFormat // Index 13 - Hyperlink style for second alternating color
                    {
                        FontId = 3, // Use the hyperlink font we defined
                        FillId = 5, // Second alternating color fill
                        BorderId = 1,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyBorder = true,
                        Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
                    }   
                )
            );
        }

        private void CreateHeaderRow(SheetData sheetData, List<(string Name, int Width, string Property)> columnDefinitions)
        {
            Row headerRow = new Row { RowIndex = 1 };
            sheetData.Append(headerRow);

            // Add cells to header row
            for (int i = 0; i < columnDefinitions.Count; i++)
            {
                string columnLetter = GetColumnLetter(i + 1);
                Cell cell = new Cell
                {
                    CellReference = $"{columnLetter}1",
                    CellValue = new CellValue(columnDefinitions[i].Name),
                    DataType = CellValues.String,
                    StyleIndex = 1 // Header style
                };
                headerRow.Append(cell);
            }
        }

        private void AddDataRows(SheetData sheetData, MergeCells mergeCells, 
                       List<TestSuiteModel> testSuites, 
                       List<(string Name, int Width, string Property)> columnDefinitions,
                       ref uint rowIndex, WorksheetPart worksheetPart)
        {
            bool useAlternateColor = false;
            uint suiteTitleStyleIndex = 2;
            uint dataStyleIndex1 = 6;
            uint dataStyleIndex2 = 7;
            uint dateStyleIndex1 = 8;
            uint dateStyleIndex2 = 9;
            uint numberStyleIndex1 = 10;
            uint numberStyleIndex2 = 11;
            _currentWorksheetPart = worksheetPart;

            foreach (var testSuite in testSuites)
            {
                // Add suite title row
                Row suiteRow = new Row { RowIndex = rowIndex++ };
                sheetData.Append(suiteRow);
                
                // Create suite title cell
                Cell suiteCell = new Cell
                {
                    CellReference = $"{GetColumnLetter(1)}{suiteRow.RowIndex}",
                    CellValue = new CellValue($"Suite: {testSuite.SuiteName}"),
                    DataType = CellValues.String,
                    StyleIndex = suiteTitleStyleIndex
                };
                suiteRow.Append(suiteCell);

                // Merge cells for suite title across all columns
                mergeCells.Append(new MergeCell
                {
                    Reference = new StringValue(
                        $"{GetColumnLetter(1)}{suiteRow.RowIndex}:{GetColumnLetter(columnDefinitions.Count)}{suiteRow.RowIndex}"
                    )
                });

                // Add test cases
                foreach (var testCase in testSuite.TestCases)
                {
                    // Alternate background color for each test case
                    useAlternateColor = !useAlternateColor;
                    uint currentDataStyleIndex = useAlternateColor ? dataStyleIndex1 : dataStyleIndex2;
                    uint currentDateStyleIndex = useAlternateColor ? dateStyleIndex1 : dateStyleIndex2;
                    uint currentNumberStyleIndex = useAlternateColor ? numberStyleIndex1 : numberStyleIndex2;

                    // Process test steps if any
                    bool isFirstStep = true;
                    if (testCase.TestSteps != null && testCase.TestSteps.Any())
                    {
                        foreach (var step in testCase.TestSteps)
                        {
                            Row row = new Row { RowIndex = rowIndex++, OutlineLevel = 1 };
                            sheetData.Append(row);

                            // Create cells for each column
                            for (int i = 0; i < columnDefinitions.Count; i++)
                            {
                                string colLetter = GetColumnLetter(i + 1);
                                string cellRef = $"{colLetter}{row.RowIndex}";
                                string property = columnDefinitions[i].Property;
                                
                                // Add appropriate cell based on property and whether it's the first step
                                if (isFirstStep)
                                {
                                    AddTestCaseCellWithSteps(row, property, cellRef, testCase, step, 
                                                            currentDataStyleIndex, currentDateStyleIndex, currentNumberStyleIndex);
                                }
                                else
                                {
                                    // For subsequent steps, only show step data
                                    AddStepOnlyCell(row, property, cellRef, step, currentDataStyleIndex);
                                }
                            }
                            
                            isFirstStep = false;
                        }
                    }
                    else
                    {
                        // No steps, create a single row for the test case
                        Row row = new Row { RowIndex = rowIndex++ };
                        sheetData.Append(row);
                        
                        // Add cells for each column
                        for (int i = 0; i < columnDefinitions.Count; i++)
                        {
                            string colLetter = GetColumnLetter(i + 1);
                            string cellRef = $"{colLetter}{row.RowIndex}";
                            string property = columnDefinitions[i].Property;
                            
                            AddTestCaseCell(row, property, cellRef, testCase, 
                                        currentDataStyleIndex, currentDateStyleIndex, currentNumberStyleIndex);
                        }
                    }
                }
            }
        }


        private void AddTestCaseCell(Row row, string property, string cellRef, TestCaseModel testCase,
                            uint dataStyleIndex, uint dateStyleIndex, uint numberStyleIndex)
        {
            // Handle step-independent properties only
            AddTestCasePropertyCell(row, property, cellRef, testCase, null,
                                   dataStyleIndex, dateStyleIndex, numberStyleIndex);
        }

        private void AddTestCaseCellWithSteps(Row row, string property, string cellRef,
                                        TestCaseModel testCase, TestStepModel step,
                                        uint dataStyleIndex, uint dateStyleIndex, uint numberStyleIndex)
        {
            // If it's a step-specific property, handle it directly
            if (property == "StepNo")
                row.Append(CreateTextCell(cellRef, step.StepNo, dataStyleIndex));
            else if (property == "StepAction")
                row.Append(CreateTextCell(cellRef, StripHtmlAndTruncate(step.StepAction), dataStyleIndex));
            else if (property == "StepExpected")
                row.Append(CreateTextCell(cellRef, StripHtmlAndTruncate(step.StepExpected), dataStyleIndex));
            else if (property == "StepRunStatus")
                row.Append(CreateTextCell(cellRef, step.StepRunStatus, dataStyleIndex));
            else if (property == "StepErrorMessage")
                row.Append(CreateTextCell(cellRef, step.StepErrorMessage, dataStyleIndex));
            else
                // For all test case properties, use the common method
                AddTestCasePropertyCell(row, property, cellRef, testCase, step,
                                       dataStyleIndex, dateStyleIndex, numberStyleIndex);
        }

        private void AddTestCasePropertyCell(Row row, string property, string cellRef,
                                           TestCaseModel testCase, TestStepModel step,
                                           uint dataStyleIndex, uint dateStyleIndex, uint numberStyleIndex)
        {
            if (property == "TestCaseId")
            {
                if (!string.IsNullOrEmpty(testCase.TestCaseUrl))
                    row.Append(CreateHyperlinkCell(cellRef, testCase.TestCaseId.ToString(), testCase.TestCaseUrl,
                             numberStyleIndex, $"Open test case {testCase.TestCaseId} in Azure Devops"));
                else
                    row.Append(CreateNumberCell(cellRef, testCase.TestCaseId.ToString(), numberStyleIndex));
            }
            else if (property == "TestCaseName")
                row.Append(CreateTextCell(cellRef, testCase.TestCaseName, dataStyleIndex));
            else if (property == "ExecutionDate" && !string.IsNullOrEmpty(testCase.ExecutionDate))
                row.Append(CreateDateCell(cellRef, DateTime.Parse(testCase.ExecutionDate), dateStyleIndex));
            else if (property == "TestCaseResult")
            {
                if (testCase.TestCaseResult != null && !string.IsNullOrEmpty(testCase.TestCaseResult.Url))
                    row.Append(CreateHyperlinkCell(cellRef, testCase.TestCaseResult.ResultMessage,
                             testCase.TestCaseResult.Url, dataStyleIndex,
                             $"Open last run result for test case {testCase.TestCaseId} in Azure Devops"));
                else if (testCase.TestCaseResult != null)
                    row.Append(CreateTextCell(cellRef, testCase.TestCaseResult.ResultMessage, dataStyleIndex));
                else
                    row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
            }
            else if (property == "FailureType")
                row.Append(CreateTextCell(cellRef, testCase.FailureType, dataStyleIndex));
            else if (property == "TestCaseComment")
                row.Append(CreateTextCell(cellRef, testCase.Comment, dataStyleIndex));
            else if (property == "RunBy")
                row.Append(CreateTextCell(cellRef, testCase.RunBy, dataStyleIndex));
            else if (property == "Configuration")
                row.Append(CreateTextCell(cellRef, testCase.Configuration, dataStyleIndex));
            else if (property == "AutomationStatus")
                row.Append(CreateTextCell(cellRef, testCase.AutomationStatus, dataStyleIndex));
            else if(property == "AssignedTo")
                row.Append(CreateTextCell(cellRef, testCase.AssignedTo, dataStyleIndex));
            else if(property == "SubSystem")
                row.Append(CreateTextCell(cellRef, testCase.SubSystem, dataStyleIndex));
            else if (property == "Priority")
                row.Append(CreateTextCell(cellRef, testCase.Priority?.ToString(), dataStyleIndex));
            else if (property == "AssociatedRequirementCount")
                row.Append(CreateTextCell(cellRef, testCase.AssociatedRequirements?.Count.ToString(), dateStyleIndex));
            else if (property.StartsWith("AssociatedRequirement_"))
                HandleRequirementCell(row, property, cellRef, testCase, dataStyleIndex);
            else if (property == "AssociatedBugCount")
                row.Append(CreateTextCell(cellRef, testCase.AssociatedBugs?.Count.ToString(), dateStyleIndex));
            else if (property.StartsWith("AssociatedBug_"))
                HandleBugCell(row, property, cellRef, testCase, dataStyleIndex);
            else if (property == "AssociatedCRCount")
                row.Append(CreateTextCell(cellRef, testCase.AssociatedCRs?.Count.ToString(), dateStyleIndex));
            else if(property.StartsWith("AssociatedCR_"))
                HandleCRCell(row, property, cellRef, testCase, dataStyleIndex);
            else
                row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
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
                        row.Append(CreateHyperlinkCell(cellRef, $"{req.Id} {req.RequirementTitle}",
                                req.Url, dataStyleIndex, $"Open Requirement {req.Id} in Azure DevOps"));
                    else
                        row.Append(CreateTextCell(cellRef, $"{req.Id} {req.RequirementTitle}", dataStyleIndex));
                }
                else
                {
                    row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
                }
            }
            else
            {
                row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
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
                        row.Append(CreateHyperlinkCell(cellRef, $"{bug.Id} {bug.BugTitle}",
                                bug.Url, dataStyleIndex, $"Open Bug {bug.Id} in Azure DevOps"));
                    else
                        row.Append(CreateTextCell(cellRef, $"{bug.Id} {bug.BugTitle}", dataStyleIndex));
                }
                else
                {
                    row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
                }
            }
            else
            {
                row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
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
                        row.Append(CreateHyperlinkCell(cellRef, $"{cr.Id} {cr.crTitle}",
                                cr.Url, dataStyleIndex, $"Open Change Request {cr.Id} in Azure DevOps"));
                    else
                        row.Append(CreateTextCell(cellRef, $"{cr.Id} {cr.crTitle}", dataStyleIndex));
                }
                else
                {
                    row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
                }
            }
            else
            {
                row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
            }
        }


        private Cell CreateHyperlinkCell(string cellReference, string displayText, string url, uint styleIndex, string tooltipMessage)
        {
            // Create a hyperlink relationship in the worksheet
            Uri uri = new Uri(url, UriKind.Absolute);

            // We need to find the current worksheet to add the hyperlink relationship
            var worksheetPart = _currentWorksheetPart;
            
            // Ensure the Worksheet object exists
            if (worksheetPart.Worksheet == null)
                worksheetPart.Worksheet = new Worksheet();

            // Create the hyperlink relationship
            HyperlinkRelationship hyperlinkRelationship = worksheetPart.AddHyperlinkRelationship(uri, true);

            // Map the original style to the corresponding hyperlink style
            uint hyperlinkStyleIndex;
            if (styleIndex == 6 || styleIndex == 8 || styleIndex == 10) // First alternating color (text, date, number)
                hyperlinkStyleIndex = 12;  // Use hyperlink style for first alternating color
            else if (styleIndex == 7 || styleIndex == 9 || styleIndex == 11) // Second alternating color (text, date, number)
                hyperlinkStyleIndex = 13;  // Use hyperlink style for second alternating color
            else
                hyperlinkStyleIndex = 12;  // Default to first hyperlink style
            
            // Create the cell with the hyperlink text and proper style
            Cell cell = new Cell
            {
                CellReference = cellReference,
                DataType = CellValues.String, // Always use string for hyperlinks
                StyleIndex = hyperlinkStyleIndex,
                CellValue = new CellValue(displayText)
            };

            // Create the hyperlink and add it to the worksheet
            Hyperlink hyperlink = new Hyperlink
            {
                Reference = cellReference,
                Id = hyperlinkRelationship.Id,
                Tooltip = tooltipMessage
            };

            // Find or create Hyperlinks element in the worksheet
            Hyperlinks hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null)
            {
                hyperlinks = new Hyperlinks();
                worksheetPart.Worksheet.Append(hyperlinks);
            }

            hyperlinks.Append(hyperlink);
            
            return cell;
        }

        private void AddStepOnlyCell(Row row, string property, string cellRef, TestStepModel step, uint dataStyleIndex)
        {
            if (property == "StepNo")
                row.Append(CreateTextCell(cellRef, step.StepNo, dataStyleIndex));
            else if (property == "StepAction")
                row.Append(CreateTextCell(cellRef, StripHtmlAndTruncate(step.StepAction), dataStyleIndex));
            else if (property == "StepExpected")
                row.Append(CreateTextCell(cellRef, StripHtmlAndTruncate(step.StepExpected), dataStyleIndex));
            else if (property == "StepRunStatus")
                row.Append(CreateTextCell(cellRef, step.StepRunStatus, dataStyleIndex));
            else if (property == "StepErrorMessage")
                row.Append(CreateTextCell(cellRef, step.StepErrorMessage, dataStyleIndex));
            else
                row.Append(CreateTextCell(cellRef, "", dataStyleIndex));
        }


        private Cell CreateTextCell(string cellReference, string cellValue, uint styleIndex = 0)
        {
            return new Cell
            {
                CellReference = cellReference,
                CellValue = new CellValue(cellValue ?? ""),
                DataType = CellValues.String,
                StyleIndex = styleIndex
            };
        }

        private Cell CreateNumberCell(string cellReference, string cellValue, uint styleIndex = 0)
        {
            return new Cell
            {
                CellReference = cellReference,
                CellValue = new CellValue(cellValue),
                DataType = CellValues.Number,
                StyleIndex = styleIndex
            };
        }

        private Cell CreateDateCell(string cellReference, DateTime date, uint styleIndex = 0)
        {
            return new Cell
            {
                CellReference = cellReference,
                CellValue = new CellValue(date.ToOADate().ToString(CultureInfo.InvariantCulture)),
                StyleIndex = styleIndex,
                DataType = CellValues.Number
            };
        }

        private string GetColumnLetter(int columnIndex)
        {
            string columnName = string.Empty;
            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                columnName = Convert.ToChar('A' + remainder) + columnName;
                columnIndex = (columnIndex - remainder - 1) / 26;
            }
            return columnName;
        }
        #endregion
    }
}
