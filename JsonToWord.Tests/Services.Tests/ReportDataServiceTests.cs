using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.ExcelServices;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class ReportDataServiceTests
    {
        [Fact]
        public void AddDataRows_ThrowsWhenSheetDataMissing()
        {
            var service = CreateService();
            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            var worksheetPart = CreateWorksheet(document);
            uint rowIndex = 1;

            Assert.Throws<ArgumentNullException>(() =>
                service.AddDataRows(null, new MergeCells(), new List<TestSuiteModel>(), new List<ColumnDefinition>(), new Dictionary<string, int>(), ref rowIndex, worksheetPart, false));
        }

        [Fact]
        public void AddDataRows_GroupBySuite_WritesRowsAndMerges()
        {
            var service = CreateService();
            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            var worksheetPart = CreateWorksheet(document);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var mergeCells = new MergeCells();

            var testSuites = BuildTestSuitesWithSteps();
            var columnDefinitions = BuildFullColumnDefinitions();
            var rowIndex = 1u;

            service.AddDataRows(sheetData, mergeCells, testSuites, columnDefinitions, new Dictionary<string, int>(), ref rowIndex, worksheetPart, true);

            Assert.True(sheetData.Elements<Row>().Count() >= 3);
            Assert.True(mergeCells.Elements<MergeCell>().Any());
            Assert.True(worksheetPart.Worksheet.Elements<Hyperlinks>().Any());
            Assert.Contains(sheetData.Descendants<Cell>(), c => c.CellValue?.Text?.Contains("Suite:") == true);
            Assert.Equal(4u, rowIndex);
        }

        [Fact]
        public void AddDataRows_NoSteps_MergesStepColumns()
        {
            var service = CreateService();
            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            var worksheetPart = CreateWorksheet(document);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var mergeCells = new MergeCells();

            var testSuites = new List<TestSuiteModel>
            {
                new TestSuiteModel
                {
                    SuiteName = "Suite",
                    TestCases = new List<TestCaseModel>
                    {
                        new TestCaseModel
                        {
                            TestCaseId = 10,
                            TestCaseName = "No Steps",
                            AssociatedRequirements = new List<AssociatedItemModel>
                            {
                                new AssociatedItemModel { Id = "R1", Title = "Req1" },
                                new AssociatedItemModel { Id = "R2", Title = "Req2" }
                            }
                        }
                    }
                }
            };

            var columnDefinitions = new List<ColumnDefinition>
            {
                Col("Step No", "StepNo", "Test Cases"),
                Col("Step Action", "StepAction", "Test Cases"),
                Col("Requirement Id", "RequirementId", "Requirements")
            };

            var rowIndex = 1u;

            service.AddDataRows(sheetData, mergeCells, testSuites, columnDefinitions, new Dictionary<string, int>(), ref rowIndex, worksheetPart, false);

            Assert.Contains(mergeCells.Elements<MergeCell>(), m => m.Reference?.Value == "A1:A2");
            Assert.Contains(mergeCells.Elements<MergeCell>(), m => m.Reference?.Value == "B1:B2");
        }

        private static ReportDataService CreateService()
        {
            var logger = new Mock<ILogger<ReportDataService>>();
            var stylesheet = new Mock<IStylesheetService>();
            return new ReportDataService(logger.Object, stylesheet.Object, new SpreadsheetService(), new ExcelHelperService());
        }

        private static WorksheetPart CreateWorksheet(SpreadsheetDocument document)
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            workbookPart.Workbook.AppendChild(new Sheets());
            workbookPart.Workbook.Sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            });

            return worksheetPart;
        }

        private static List<TestSuiteModel> BuildTestSuitesWithSteps()
        {
            return new List<TestSuiteModel>
            {
                new TestSuiteModel
                {
                    SuiteName = "Suite 1",
                    TestCases = new List<TestCaseModel>
                    {
                        new TestCaseModel
                        {
                            TestCaseId = 1,
                            TestCaseName = "Case 1",
                            TestCaseUrl = "https://example.com/tc1",
                            ExecutionDate = "2024-01-01",
                            TestCaseResult = new TestCaseResultModel
                            {
                                ResultMessage = "Passed",
                                Url = "https://example.com/result"
                            },
                            FailureType = "None",
                            Comment = "Comment",
                            RunBy = "User",
                            Configuration = "Config",
                            State = "Active",
                            StateChangeDate = "2024-01-02",
                            TestSteps = new List<TestStepModel>
                            {
                                new TestStepModel
                                {
                                    StepNo = "1",
                                    StepAction = "<b>Do</b>",
                                    StepExpected = "Expect",
                                    StepRunStatus = "Pass",
                                    StepErrorMessage = "Err"
                                },
                                new TestStepModel
                                {
                                    StepNo = "2",
                                    StepAction = "Do2",
                                    StepExpected = "Expect2"
                                }
                            },
                            HistoryEntries = new List<string> { "<p>History</p>", "History2" },
                            AssociatedRequirements = new List<AssociatedItemModel>
                            {
                                new AssociatedItemModel
                                {
                                    Id = "R1",
                                    Title = "Req1",
                                    Url = "https://example.com/req1",
                                    CustomFields = new Dictionary<string, object>
                                    {
                                        { "customDate", "2024-02-01" }
                                    }
                                }
                            },
                            AssociatedBugs = new List<AssociatedItemModel>
                            {
                                new AssociatedItemModel
                                {
                                    Id = "B1",
                                    Title = "Bug1"
                                }
                            },
                            AssociatedCRs = new List<AssociatedItemModel>
                            {
                                new AssociatedItemModel
                                {
                                    Id = "C1",
                                    Title = "CR1",
                                    Url = "https://example.com/cr1"
                                }
                            },
                            CustomFields = new Dictionary<string, object>
                            {
                                { "customDate", "2024-03-01" },
                                { "customText", "Value" }
                            }
                        }
                    }
                }
            };
        }

        private static List<ColumnDefinition> BuildFullColumnDefinitions()
        {
            return new List<ColumnDefinition>
            {
                Col("Suite", "SuiteName", "Test Cases"),
                Col("Id", "TestCaseId", "Test Cases"),
                Col("Name", "TestCaseName", "Test Cases"),
                Col("Execution Date", "ExecutionDate", "Test Cases"),
                Col("Result", "TestCaseResult", "Test Cases"),
                Col("Failure Type", "FailureType", "Test Cases"),
                Col("Comment", "TestCaseComment", "Test Cases"),
                Col("Run By", "RunBy", "Test Cases"),
                Col("Configuration", "Configuration", "Test Cases"),
                Col("State", "State", "Test Cases"),
                Col("State Change", "StateChangeDate", "Test Cases"),
                Col("Req Count", "AssociatedRequirementCount", "Test Cases"),
                Col("Req 0", "AssociatedRequirement_0", "Test Cases"),
                Col("Bug Count", "AssociatedBugCount", "Test Cases"),
                Col("Bug 0", "AssociatedBug_0", "Test Cases"),
                Col("CR Count", "AssociatedCRCount", "Test Cases"),
                Col("CR 0", "AssociatedCR_0", "Test Cases"),
                Col("Custom Date", "CustomDate", "Test Cases"),
                Col("Custom Text", "CustomText", "Test Cases"),
                Col("Step No", "StepNo", "Test Cases"),
                Col("Step Action", "StepAction", "Test Cases"),
                Col("Step Expected", "StepExpected", "Test Cases"),
                Col("Step Status", "StepRunStatus", "Test Cases"),
                Col("Step Error", "StepErrorMessage", "Test Cases"),
                Col("History", "History", "Test Cases"),
                Col("Req Id", "RequirementId", "Requirements"),
                Col("Req Name", "RequirementName", "Requirements"),
                Col("Req Custom Date", "CustomDate", "Requirements"),
                Col("Bug Id", "BugId", "Bugs"),
                Col("Bug Name", "BugName", "Bugs"),
                Col("CR Id", "CRId", "CRs"),
                Col("CR Name", "CRName", "CRs")
            };
        }

        private static ColumnDefinition Col(string name, string property, string group)
        {
            return new ColumnDefinition { Name = name, Property = property, Group = group };
        }

    }
}
