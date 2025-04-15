using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Tests.Services
{
    public class TestReporterServiceTest : IDisposable
    {
        private readonly Mock<ILogger<TestReporterService>> _loggerMock;
        private readonly TestReporterService _testReporterService;
        private readonly string _tempFilePath;

        public TestReporterServiceTest()
        {
            _loggerMock = new Mock<ILogger<TestReporterService>>();
            _testReporterService = new TestReporterService(_loggerMock.Object);
            _tempFilePath = Path.Combine(Path.GetTempPath(), $"TestReport_{Guid.NewGuid()}.xlsx");
        }

        public void Dispose()
        {
            if (File.Exists(_tempFilePath))
            {
                try
                {
                    File.Delete(_tempFilePath);
                }
                catch (Exception)
                {
                    // Ignore exceptions during cleanup
                }
            }
        }

        [Fact]
        public void Insert_WithNullDocument_ThrowsArgumentNullException()
        {
            // Arrange
            SpreadsheetDocument document = null;
            var testReporterModel = CreateSampleTestReporterModel();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _testReporterService.Insert(document, "Sheet1", testReporterModel));
        }

        [Fact]
        public void Insert_WithEmptyWorksheetName_ThrowsArgumentException()
        {
            // Arrange
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateSampleTestReporterModel();

            // Act & Assert
            Assert.Throws<ArgumentException>(() => _testReporterService.Insert(document, "", testReporterModel));
        }

        [Fact]
        public void Insert_WithNullTestReporterModel_ThrowsArgumentNullException()
        {
            // Arrange
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            TestReporterModel testReporterModel = null;

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _testReporterService.Insert(document, "Sheet1", testReporterModel));
        }

        [Fact]
        public void Insert_CreatesWorksheetWithCorrectName()
        {
            // Arrange
            const string worksheetName = "Test Worksheet";
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateSampleTestReporterModel();

            // Act
            _testReporterService.Insert(document, worksheetName, testReporterModel);
            document.Dispose();

            // Assert
            using var openedDoc = SpreadsheetDocument.Open(_tempFilePath, false);
            var workbookPart = openedDoc.WorkbookPart;
            Assert.NotNull(workbookPart);

            var sheets = workbookPart.Workbook.Sheets;
            Assert.NotNull(sheets);

            var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
            Assert.NotNull(sheet);
        }

        [Fact]
        public void Insert_CreatesHeaderRowWithCorrectColumns()
        {
            // Arrange
            const string worksheetName = "Test Worksheet";
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateSampleTestReporterModel();

            // Act
            _testReporterService.Insert(document, worksheetName, testReporterModel);
            document.Dispose();

            // Assert
            using var openedDoc = SpreadsheetDocument.Open(_tempFilePath, false);
            var workbookPart = openedDoc.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // Get the header row (first row)
            var headerRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == 1);
            Assert.NotNull(headerRow);

            // Check that we have expected headers
            var headerCells = headerRow.Elements<Cell>().ToList();
            Assert.Contains(headerCells, c => GetCellValue(c, workbookPart) == "Test Case ID");
            Assert.Contains(headerCells, c => GetCellValue(c, workbookPart) == "Test Case Title");
            Assert.Contains(headerCells, c => GetCellValue(c, workbookPart) == "Execution Date");
        }

        [Fact]
        public void Insert_CreatesTestSuiteTitleRow()
        {
            // Arrange
            const string worksheetName = "Test Worksheet";
            const string suiteName = "Sample Test Suite";
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateSampleTestReporterModel(suiteName);

            // Act
            _testReporterService.Insert(document, worksheetName, testReporterModel);
            document.Dispose();

            // Assert
            using var openedDoc = SpreadsheetDocument.Open(_tempFilePath, false);
            var workbookPart = openedDoc.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // The suite title should be on the second row
            var suiteRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == 2);
            Assert.NotNull(suiteRow);

            // Check that the suite title contains the suite name
            var firstCell = suiteRow.Elements<Cell>().FirstOrDefault();
            Assert.NotNull(firstCell);
            var cellValue = GetCellValue(firstCell, workbookPart);
            Assert.Contains(suiteName, cellValue);
        }

        [Fact]
        public void Insert_TestCaseWithSteps_CreatesMultipleRows()
        {
            // Arrange
            const string worksheetName = "Test Worksheet";
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateTestReporterModelWithSteps();

            // Act
            _testReporterService.Insert(document, worksheetName, testReporterModel);
            document.Dispose();

            // Assert
            using var openedDoc = SpreadsheetDocument.Open(_tempFilePath, false);
            var workbookPart = openedDoc.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // Count the rows - should be 1 (header) + 1 (suite) + 2 (two steps) = 4
            var rows = sheetData.Elements<Row>().ToList();
            Assert.Equal(4, rows.Count);

            // Check that the step data is present in rows
            var stepRows = rows.Where(r => r.RowIndex > 2).ToList();
            Assert.Equal(2, stepRows.Count);

            // Check that the step actions from our test data are present
            bool foundStep1 = false;
            bool foundStep2 = false;

            foreach (var row in stepRows)
            {
                var cells = row.Elements<Cell>().ToList();
                foreach (var cell in cells)
                {
                    var value = GetCellValue(cell, workbookPart);
                    if (value == "Step 1 Action") foundStep1 = true;
                    if (value == "Step 2 Action") foundStep2 = true;
                }
            }

            Assert.True(foundStep1, "Step 1 action not found in cells");
            Assert.True(foundStep2, "Step 2 action not found in cells");
        }

        [Fact]
        public void Insert_WithHyperlink_CreatesHyperlinkInWorksheet()
        {
            // Arrange
            const string worksheetName = "Test Worksheet";
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateTestReporterModelWithHyperlink();

            // Act
            _testReporterService.Insert(document, worksheetName, testReporterModel);
            document.Dispose();

            // Assert
            using var openedDoc = SpreadsheetDocument.Open(_tempFilePath, false);
            var workbookPart = openedDoc.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

            // Check for hyperlinks
            var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
            Assert.NotNull(hyperlinks);
            Assert.True(hyperlinks.Count() > 0, "No hyperlinks found in worksheet");

            // Verify at least one hyperlink has the URL we expect
            bool foundHyperlink = false;
            foreach (var hyperlink in hyperlinks.Elements<Hyperlink>())
            {
                var relationship = worksheetPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id);
                if (relationship != null && relationship.Uri.ToString().Contains("example.com"))
                {
                    foundHyperlink = true;
                    break;
                }
            }

            Assert.True(foundHyperlink, "Expected hyperlink not found");
        }

        [Fact]
        public void Insert_AddsStylesheetToWorkbook()
        {
            // Arrange
            const string worksheetName = "Test Worksheet";
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateSampleTestReporterModel();

            // Act
            _testReporterService.Insert(document, worksheetName, testReporterModel);
            document.Dispose();

            // Assert
            using var openedDoc = SpreadsheetDocument.Open(_tempFilePath, false);
            var workbookPart = openedDoc.WorkbookPart;

            // Check that a stylesheet part was created
            var stylesPart = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            Assert.NotNull(stylesPart);

            // Check that the stylesheet contains expected elements
            var stylesheet = stylesPart.Stylesheet;
            Assert.NotNull(stylesheet);
            Assert.NotNull(stylesheet.Fonts);
            Assert.NotNull(stylesheet.Fills);
            Assert.NotNull(stylesheet.Borders);
            Assert.NotNull(stylesheet.CellFormats);
        }

        [Fact]
        public void Insert_WithAssociatedRequirements_CreatesRequirementColumns()
        {
            // Arrange
            const string worksheetName = "Test Worksheet";
            using var document = SpreadsheetDocument.Create(_tempFilePath, SpreadsheetDocumentType.Workbook);
            var testReporterModel = CreateTestReporterModelWithRequirements();

            // Act
            _testReporterService.Insert(document, worksheetName, testReporterModel);
            document.Dispose();

            // Assert
            using var openedDoc = SpreadsheetDocument.Open(_tempFilePath, false);
            var workbookPart = openedDoc.WorkbookPart;
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // Get the header row
            var headerRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == 1);
            Assert.NotNull(headerRow);

            // Check for requirement columns in headers
            var headerCells = headerRow.Elements<Cell>().ToList();
            bool foundReqCountHeader = false;
            bool foundReqHeader = false;

            foreach (var cell in headerCells)
            {
                var value = GetCellValue(cell, workbookPart);
                if (value == "Associated Req. Count") foundReqCountHeader = true;
                if (value.StartsWith("Associated Req.")) foundReqHeader = true;
            }

            Assert.True(foundReqCountHeader, "Associated Requirement Count column not found");
            Assert.True(foundReqHeader, "Associated Requirement column not found");
        }

        #region Helper Methods

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            // Handle null cases
            if (cell == null || cell.CellValue == null) return string.Empty;

            string value = cell.CellValue.InnerText;

            // If this is a shared string, look it up
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTable != null)
                {
                    int index = int.Parse(value);
                    if (index < stringTable.SharedStringTable.Elements<SharedStringItem>().Count())
                    {
                        value = stringTable.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
                    }
                }
            }

            return value;
        }

        private TestReporterModel CreateSampleTestReporterModel(string suiteName = "Sample Test Suite")
        {
            return new TestReporterModel
            {
                TestPlanName = "Test Plan",
                TestSuites = new System.Collections.Generic.List<TestSuiteModel>
                {
                    new TestSuiteModel
                    {
                        SuiteName = suiteName,
                        TestCases = new System.Collections.Generic.List<TestCaseModel>
                        {
                            new TestCaseModel
                            {
                                TestCaseId = 12345,
                                TestCaseName = "Sample Test Case",
                                ExecutionDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                TestCaseResult = new TestCaseResultModel
                                {
                                    ResultMessage = "Passed"
                                },
                                RunBy = "Test User"
                            }
                        }
                    }
                }
            };
        }

        private TestReporterModel CreateTestReporterModelWithSteps()
        {
            var model = CreateSampleTestReporterModel();

            // Add test steps to the first test case
            model.TestSuites[0].TestCases[0].TestSteps = new System.Collections.Generic.List<TestStepModel>
            {
                new TestStepModel
                {
                    StepNo = "1",
                    StepAction = "Step 1 Action",
                    StepExpected = "Step 1 Expected Result",
                    StepRunStatus = "Passed"
                },
                new TestStepModel
                {
                    StepNo = "2",
                    StepAction = "Step 2 Action",
                    StepExpected = "Step 2 Expected Result",
                    StepRunStatus = "Passed"
                }
            };

            return model;
        }

        private TestReporterModel CreateTestReporterModelWithHyperlink()
        {
            var model = CreateSampleTestReporterModel();

            // Add URL to test case
            model.TestSuites[0].TestCases[0].TestCaseUrl = "https://example.com/testcase/12345";

            return model;
        }

        private TestReporterModel CreateTestReporterModelWithRequirements()
        {
            var model = CreateSampleTestReporterModel();

            // Add associated requirements
            model.TestSuites[0].TestCases[0].AssociatedRequirements = new List<AssociatedRequirementModel>
            {
                new AssociatedRequirementModel
                {
                    Id = "REQ-001",
                    RequirementTitle = "Sample Requirement",
                    Url = "https://example.com/requirement/REQ-001"
                }
            };

            return model;
        }

        #endregion
    }
}