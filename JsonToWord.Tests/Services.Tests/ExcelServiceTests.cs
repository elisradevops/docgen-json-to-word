using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class ExcelServiceTests
    {
        [Fact]
        public void CreateExcelDocument_ThrowsAndCleansUp_WhenNoWordObjects()
        {
            var logger = new Mock<ILogger<ExcelService>>();
            var testReporterService = new Mock<ITestReporterService>();
            var service = new ExcelService(logger.Object, testReporterService.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, "report.xlsx");

            var model = new ExcelModel
            {
                LocalPath = filePath,
                ContentControls = new List<TestReporterContentControl>
                {
                    new TestReporterContentControl
                    {
                        Title = "cc",
                        WordObjects = new List<ITestReporterObject>()
                    }
                }
            };

            try
            {
                Assert.Throws<Exception>(() => service.CreateExcelDocument(model));
                Assert.False(File.Exists(filePath));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateExcelDocument_CallsTestReporterService()
        {
            var logger = new Mock<ILogger<ExcelService>>();
            var testReporterService = new Mock<ITestReporterService>();
            var service = new ExcelService(logger.Object, testReporterService.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, "report.xlsx");

            var testReporter = new TestReporterModel
            {
                TestPlanName = "Plan A"
            };

            testReporterService
                .Setup(s => s.Insert(It.IsAny<SpreadsheetDocument>(), It.IsAny<string>(), It.IsAny<TestReporterModel>(), It.IsAny<bool>()))
                .Callback<SpreadsheetDocument, string, TestReporterModel, bool>((document, _, __, ___) =>
                {
                    if (document.WorkbookPart == null)
                    {
                        var workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                    }
                });

            var model = new ExcelModel
            {
                LocalPath = filePath,
                ContentControls = new List<TestReporterContentControl>
                {
                    new TestReporterContentControl
                    {
                        Title = "cc",
                        WordObjects = new List<ITestReporterObject> { testReporter }
                    }
                }
            };

            try
            {
                var resultPath = service.CreateExcelDocument(model);

                Assert.Equal(filePath, resultPath);
                Assert.True(File.Exists(filePath));
                testReporterService.Verify(s => s.Insert(It.IsAny<DocumentFormat.OpenXml.Packaging.SpreadsheetDocument>(), "Plan A", testReporter, It.IsAny<bool>()), Times.Once);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }
    }
}
