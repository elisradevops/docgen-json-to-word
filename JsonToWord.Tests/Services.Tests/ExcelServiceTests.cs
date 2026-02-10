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
            var flatTestReporterService = new Mock<IFlatTestReporterService>();
            var mewpCoverageReporterService = new Mock<IMewpCoverageReporterService>();
            var service = new ExcelService(
                logger.Object,
                testReporterService.Object,
                flatTestReporterService.Object,
                mewpCoverageReporterService.Object
            );

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
            var flatTestReporterService = new Mock<IFlatTestReporterService>();
            var mewpCoverageReporterService = new Mock<IMewpCoverageReporterService>();
            var service = new ExcelService(
                logger.Object,
                testReporterService.Object,
                flatTestReporterService.Object,
                mewpCoverageReporterService.Object
            );

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

        [Fact]
        public void CreateExcelDocument_CallsFlatTestReporterService()
        {
            var logger = new Mock<ILogger<ExcelService>>();
            var testReporterService = new Mock<ITestReporterService>();
            var flatTestReporterService = new Mock<IFlatTestReporterService>();
            var mewpCoverageReporterService = new Mock<IMewpCoverageReporterService>();
            var service = new ExcelService(
                logger.Object,
                testReporterService.Object,
                flatTestReporterService.Object,
                mewpCoverageReporterService.Object
            );

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, "flat-report.xlsx");

            var flatReporter = new FlatTestReporterModel
            {
                TestPlanName = "Flat Plan",
                Rows = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object> { { "PlanID", "1" } }
                }
            };

            flatTestReporterService
                .Setup(s => s.Insert(It.IsAny<SpreadsheetDocument>(), It.IsAny<string>(), It.IsAny<FlatTestReporterModel>()))
                .Callback<SpreadsheetDocument, string, FlatTestReporterModel>((document, _, __) =>
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
                        WordObjects = new List<ITestReporterObject> { flatReporter }
                    }
                }
            };

            try
            {
                var resultPath = service.CreateExcelDocument(model);

                Assert.Equal(filePath, resultPath);
                Assert.True(File.Exists(filePath));
                flatTestReporterService.Verify(
                    s => s.Insert(It.IsAny<SpreadsheetDocument>(), "Flat Plan", flatReporter),
                    Times.Once
                );
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateExcelDocument_CallsTestReporterService_ForEachReporterWorksheet()
        {
            var logger = new Mock<ILogger<ExcelService>>();
            var testReporterService = new Mock<ITestReporterService>();
            var flatTestReporterService = new Mock<IFlatTestReporterService>();
            var mewpCoverageReporterService = new Mock<IMewpCoverageReporterService>();
            var service = new ExcelService(
                logger.Object,
                testReporterService.Object,
                flatTestReporterService.Object,
                mewpCoverageReporterService.Object
            );

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, "multi-sheet-report.xlsx");

            var mainReporter = new TestReporterModel
            {
                TestPlanName = "Plan A"
            };
            var mewpCoverageReporter = new TestReporterModel
            {
                TestPlanName = "MEWP L2 Coverage - Plan A"
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
                        WordObjects = new List<ITestReporterObject> { mainReporter, mewpCoverageReporter }
                    }
                }
            };

            try
            {
                var resultPath = service.CreateExcelDocument(model);

                Assert.Equal(filePath, resultPath);
                Assert.True(File.Exists(filePath));
                testReporterService.Verify(
                    s => s.Insert(It.IsAny<SpreadsheetDocument>(), "Plan A", mainReporter, It.IsAny<bool>()),
                    Times.Once
                );
                testReporterService.Verify(
                    s => s.Insert(It.IsAny<SpreadsheetDocument>(), "MEWP L2 Coverage - Plan A", mewpCoverageReporter, It.IsAny<bool>()),
                    Times.Once
                );
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateExcelDocument_CallsMewpCoverageReporterService()
        {
            var logger = new Mock<ILogger<ExcelService>>();
            var testReporterService = new Mock<ITestReporterService>();
            var flatTestReporterService = new Mock<IFlatTestReporterService>();
            var mewpCoverageReporterService = new Mock<IMewpCoverageReporterService>();
            var service = new ExcelService(
                logger.Object,
                testReporterService.Object,
                flatTestReporterService.Object,
                mewpCoverageReporterService.Object
            );

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, "mewp-coverage-report.xlsx");

            var mewpCoverageReporter = new MewpCoverageReporterModel
            {
                TestPlanName = "MEWP L2 Coverage - Plan A",
                Rows = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object> { { "Customer ID", "SR1001" } }
                }
            };

            mewpCoverageReporterService
                .Setup(s => s.Insert(It.IsAny<SpreadsheetDocument>(), It.IsAny<string>(), It.IsAny<MewpCoverageReporterModel>()))
                .Callback<SpreadsheetDocument, string, MewpCoverageReporterModel>((document, _, __) =>
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
                        WordObjects = new List<ITestReporterObject> { mewpCoverageReporter }
                    }
                }
            };

            try
            {
                var resultPath = service.CreateExcelDocument(model);

                Assert.Equal(filePath, resultPath);
                Assert.True(File.Exists(filePath));
                mewpCoverageReporterService.Verify(
                    s => s.Insert(It.IsAny<SpreadsheetDocument>(), "MEWP L2 Coverage - Plan A", mewpCoverageReporter),
                    Times.Once
                );
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }
    }
}
