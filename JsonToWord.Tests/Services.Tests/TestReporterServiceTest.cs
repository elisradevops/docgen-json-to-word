using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;

namespace JsonToWord.Services.Tests
{
    public class TestReporterServiceTest : IDisposable
    {
        private readonly Mock<ILogger<TestReporterService>> _mockLogger;
        private readonly Mock<IColumnService> _mockColumnService;
        private readonly Mock<ISpreadsheetService> _mockSpreadsheetService;
        private readonly Mock<IReportDataService> _mockReportDataService;
        private readonly Mock<IStylesheetService> _mockStylesheetService;
        private readonly TestReporterService _testReporterService;

        private readonly MemoryStream _memoryStream;
        private readonly SpreadsheetDocument _spreadsheetDocument;

        public TestReporterServiceTest()
        {
            _mockLogger = new Mock<ILogger<TestReporterService>>();
            _mockColumnService = new Mock<IColumnService>();
            _mockSpreadsheetService = new Mock<ISpreadsheetService>();
            _mockReportDataService = new Mock<IReportDataService>();
            _mockStylesheetService = new Mock<IStylesheetService>();

            _testReporterService = new TestReporterService(
                _mockLogger.Object,
                _mockColumnService.Object,
                _mockSpreadsheetService.Object,
                _mockReportDataService.Object,
                _mockStylesheetService.Object);

            _memoryStream = new MemoryStream();
            _spreadsheetDocument = SpreadsheetDocument.Create(_memoryStream, SpreadsheetDocumentType.Workbook);
        }

        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
            _memoryStream.Dispose();
            GC.SuppressFinalize(this);
        }

        [Fact]
        public void Insert_WithNullDocument_ThrowsArgumentNullException()
        {
            // Arrange
            var model = new TestReporterModel();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _testReporterService.Insert(null, "Sheet1", model, false));
        }

        [Fact]
        public void Insert_WithEmptyWorksheetName_ThrowsArgumentException()
        {
            // Arrange
            var model = new TestReporterModel();

            // Act & Assert
            Assert.Throws<ArgumentException>(() => _testReporterService.Insert(_spreadsheetDocument, "", model, false));
        }

        [Fact]
        public void Insert_WithNullWorksheetName_ThrowsArgumentException()
        {
            // Arrange
            var model = new TestReporterModel();

            // Act & Assert
            Assert.Throws<ArgumentException>(() => _testReporterService.Insert(_spreadsheetDocument, null, model, false));
        }

        [Fact]
        public void Insert_WithNullTestReporterModel_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _testReporterService.Insert(_spreadsheetDocument, "Sheet1", null, false));
        }

        [Fact]
        public void Insert_WhenServiceThrowsException_LogsErrorAndRethrows()
        {
            // Arrange
            var model = new TestReporterModel();
            var exception = new Exception("Test exception");
            var worksheetName = "Sheet1";

            _mockSpreadsheetService.Setup(s => s.GetOrCreateWorksheetPart(It.IsAny<WorkbookPart>(), It.IsAny<string>()))
                .Throws(exception);

            // Act & Assert
            var ex = Assert.Throws<Exception>(() => _testReporterService.Insert(_spreadsheetDocument, worksheetName, model, false));
            Assert.Equal(exception, ex);

            _mockLogger.Verify(
                x => x.Log(
                    LogLevel.Error,
                    It.IsAny<EventId>(),
                    It.Is<It.IsAnyType>((v, t) => v.ToString().Contains($"Error inserting grouped table into worksheet '{worksheetName}'")),
                    exception,
                    It.IsAny<Func<It.IsAnyType, Exception, string>>()),
                Times.Once);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Insert_WithValidInputs_CallsServicesAndSavesSuccessfully(bool groupBySuite)
        {
            // Arrange
            const string worksheetName = "Test Report";
            var model = new TestReporterModel
            {
                TestSuites = new List<TestSuiteModel>
                {
                    new TestSuiteModel
                    {
                        SuiteName = "Suite 1",
                        TestCases = new List<TestCaseModel>
                        {
                            // Test Case with multiple requirements and bugs
                            new TestCaseModel
                            {
                                TestCaseId = 1,
                                TestCaseName = "TC 1: Complex associations",
                                AssociatedRequirements = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel { Id = "R1", Title = "Requirement 1" },
                                    new AssociatedItemModel { Id = "R2", Title = "Requirement 2" }
                                },
                                AssociatedBugs = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel { Id = "B1", Title = "Bug 1" }
                                },
                                AssociatedCRs = new List<AssociatedItemModel>() // Empty list
                            },
                            // Test Case with one of each associated item
                            new TestCaseModel
                            {
                                TestCaseId = 2,
                                TestCaseName = "TC 2: One of each",
                                AssociatedRequirements = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel { Id = "R3", Title = "Requirement 3" }
                                },
                                AssociatedBugs = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel { Id = "B2", Title = "Bug 2" }
                                },
                                AssociatedCRs = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel { Id = "CR1", Title = "Change Request 1" }
                                }
                            },
                            // Test Case with no associated items
                            new TestCaseModel
                            {
                                TestCaseId = 3,
                                TestCaseName = "TC 3: No associations"
                            }
                        }
                    },
                    new TestSuiteModel
                    {
                        SuiteName = "Suite 2",
                        TestCases = new List<TestCaseModel>
                        {
                            // Test Case with only CRs
                            new TestCaseModel
                            {
                                TestCaseId = 4,
                                TestCaseName = "TC 4: Only CRs",
                                AssociatedCRs = new List<AssociatedItemModel>
                                {
                                    new AssociatedItemModel { Id = "CR2", Title = "Change Request 2" },
                                    new AssociatedItemModel { Id = "CR3", Title = "Change Request 3" }
                                }
                            }
                        }
                    }
                }
            };

            var workbookPart = _spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var columnDefinitions = new List<ColumnDefinition> { new ColumnDefinition { Name = "ID", Width = 10, Group = "Test Cases" } };
            var columnCountForeachGroup = new Dictionary<string, int> { { "Test Cases", 1 } };

            _mockSpreadsheetService.Setup(s => s.GetOrCreateWorksheetPart(It.IsAny<WorkbookPart>(), worksheetName)).Returns(worksheetPart);
            _mockColumnService.Setup(c => c.DefineColumns(model, groupBySuite)).Returns(columnDefinitions);
            _mockColumnService.Setup(c => c.CreateColumns(columnDefinitions)).Returns(new Columns());
            _mockColumnService.Setup(c => c.GetColumnCountForeachGroup(columnDefinitions)).Returns(columnCountForeachGroup);

            // Act
            _testReporterService.Insert(_spreadsheetDocument, worksheetName, model, groupBySuite);

            // Assert
            _mockSpreadsheetService.Verify(s => s.GetOrCreateWorksheetPart(It.IsAny<WorkbookPart>(), worksheetName), Times.Once);
            _mockColumnService.Verify(c => c.DefineColumns(model, groupBySuite), Times.Once);
            _mockColumnService.Verify(c => c.CreateColumns(columnDefinitions), Times.Once);
            _mockColumnService.Verify(c => c.GetColumnCountForeachGroup(columnDefinitions), Times.Once);
            _mockSpreadsheetService.Verify(s => s.CreateHeaderRow(It.IsAny<SheetData>(), columnDefinitions, It.IsAny<MergeCells>(), columnCountForeachGroup, It.IsAny<Dictionary<string, int>>()), Times.Once);

            _mockReportDataService.Verify(r => r.AddDataRows(It.IsAny<SheetData>(), It.IsAny<MergeCells>(), model.TestSuites, columnDefinitions, columnCountForeachGroup, ref It.Ref<uint>.IsAny, worksheetPart, groupBySuite), Times.Once);

            _mockStylesheetService.Verify(s => s.EnsureStylesheet(It.IsAny<WorkbookPart>()), Times.Once);
        }
    }
}
