using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services;
using JsonToWord.Services.ExcelServices;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;

namespace JsonToWord.Services.Tests
{
    public class InternalValidationReporterServiceTests
    {
        [Fact]
        public void Insert_WithNullDocument_ThrowsArgumentNullException()
        {
            var service = new InternalValidationReporterService(
                new Mock<ILogger<InternalValidationReporterService>>().Object,
                new SpreadsheetService(),
                new StylesheetService()
            );

            var model = new InternalValidationReporterModel();

            Assert.Throws<ArgumentNullException>(() => service.Insert(null, "Sheet1", model));
        }

        [Fact]
        public void Insert_WithNullModel_ThrowsArgumentNullException()
        {
            var service = new InternalValidationReporterService(
                new Mock<ILogger<InternalValidationReporterService>>().Object,
                new SpreadsheetService(),
                new StylesheetService()
            );

            using var memoryStream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);

            Assert.Throws<ArgumentNullException>(() => service.Insert(document, "Sheet1", null));
        }

        [Fact]
        public void Insert_WithBlankWorksheetName_UsesDefaultSheetAndDefaultHeaders()
        {
            var service = new InternalValidationReporterService(
                new Mock<ILogger<InternalValidationReporterService>>().Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using (var document = SpreadsheetDocument.Create(tempPath, SpreadsheetDocumentType.Workbook))
                {
                    var model = new InternalValidationReporterModel
                    {
                        Rows = new List<Dictionary<string, object>>()
                    };

                    service.Insert(document, "  ", model);

                    var workbookPart = document.WorkbookPart!;
                    var createdSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault();
                    Assert.NotNull(createdSheet);
                    Assert.Equal("MEWP Internal Validation", createdSheet!.Name!.Value);

                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(createdSheet.Id!);
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var headerRow = sheetData.Elements<Row>().First();
                    var headerValues = headerRow.Elements<Cell>().Select(c => c.CellValue!.Text).ToList();

                    Assert.Equal(
                        new[]
                        {
                            "Test Case ID",
                            "Test Case Title",
                            "Mentioned but Not Linked",
                            "Linked but Not Mentioned",
                            "Validation Status",
                        },
                        headerValues
                    );
                }
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_WritesNumericAndTextCellsWithExpectedDataTypesAndStyles()
        {
            var service = new InternalValidationReporterService(
                new Mock<ILogger<InternalValidationReporterService>>().Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using (var document = SpreadsheetDocument.Create(tempPath, SpreadsheetDocumentType.Workbook))
                {
                    var model = new InternalValidationReporterModel
                    {
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "Test Case ID", 101 },
                                { "Test Case Title", "Title A" },
                                { "Mentioned but Not Linked", "None" },
                                { "Linked but Not Mentioned", "REQ-55" },
                                { "Validation Status", "Pass" },
                            },
                            new Dictionary<string, object>
                            {
                                { "Test Case ID", "TC-ABC" },
                                { "Test Case Title", "Title B" },
                                { "Mentioned but Not Linked", "REQ-77" },
                                { "Linked but Not Mentioned", "None" },
                                { "Validation Status", "Fail" },
                            },
                        },
                    };

                    service.Insert(document, "Internal Validation", model);

                    var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var rows = sheetData.Elements<Row>().ToList();

                    var firstDataRowFirstCell = rows[1].Elements<Cell>().First();
                    Assert.Equal(CellValues.Number, firstDataRowFirstCell.DataType!.Value);
                    Assert.Equal((uint)10, firstDataRowFirstCell.StyleIndex!.Value);

                    var secondDataRowFirstCell = rows[2].Elements<Cell>().First();
                    Assert.Equal(CellValues.String, secondDataRowFirstCell.DataType!.Value);
                    Assert.Equal((uint)11, secondDataRowFirstCell.StyleIndex!.Value);
                }
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_WithCustomColumnOrder_AppliesHeaderOrderAndColumnWidths()
        {
            var service = new InternalValidationReporterService(
                new Mock<ILogger<InternalValidationReporterService>>().Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using (var document = SpreadsheetDocument.Create(tempPath, SpreadsheetDocumentType.Workbook))
                {
                    var model = new InternalValidationReporterModel
                    {
                        ColumnOrder = new List<string>
                        {
                            "Validation Status",
                            "Test Case ID",
                        },
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "Validation Status", "Pass" },
                                { "Test Case ID", 5001 },
                            },
                        },
                    };

                    service.Insert(document, "Custom Order", model);

                    var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var headerRow = sheetData.Elements<Row>().First();
                    var headers = headerRow.Elements<Cell>().Select(c => c.CellValue!.Text).ToList();
                    Assert.Equal(new[] { "Validation Status", "Test Case ID" }, headers);

                    var columns = worksheetPart.Worksheet.Elements<Columns>().First().Elements<Column>().ToList();
                    Assert.Equal(2, columns.Count);
                    Assert.Equal(20d, columns[0].Width!.Value);
                    Assert.Equal(16d, columns[1].Width!.Value);
                }
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        [Fact]
        public void Insert_WhenSpreadsheetServiceThrows_LogsErrorAndRethrows()
        {
            var logger = new Mock<ILogger<InternalValidationReporterService>>();
            var spreadsheetService = new Mock<ISpreadsheetService>();
            var stylesheetService = new Mock<IStylesheetService>();
            var service = new InternalValidationReporterService(
                logger.Object,
                spreadsheetService.Object,
                stylesheetService.Object
            );

            using var memoryStream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);

            var model = new InternalValidationReporterModel
            {
                Rows = new List<Dictionary<string, object>>(),
            };
            var exception = new InvalidOperationException("Spreadsheet failure");
            const string worksheetName = "Validation";

            spreadsheetService
                .Setup(s => s.GetOrCreateWorksheetPart(It.IsAny<WorkbookPart>(), It.IsAny<string>()))
                .Throws(exception);

            var ex = Assert.Throws<InvalidOperationException>(() => service.Insert(document, worksheetName, model));
            Assert.Equal(exception, ex);

            logger.Verify(
                x => x.Log(
                    LogLevel.Error,
                    It.IsAny<EventId>(),
                    It.Is<It.IsAnyType>((v, t) => v.ToString()!.Contains($"Error inserting Internal Validation worksheet '{worksheetName}'")),
                    exception,
                    It.IsAny<Func<It.IsAnyType, Exception?, string>>()
                ),
                Times.Once
            );
        }
    }
}
