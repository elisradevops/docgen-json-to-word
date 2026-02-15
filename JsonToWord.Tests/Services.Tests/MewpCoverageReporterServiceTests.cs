using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.TestReporterModels;
using JsonToWord.Services;
using JsonToWord.Services.ExcelServices;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class MewpCoverageReporterServiceTests
    {
        [Fact]
        public void Insert_WritesNumericColumnsAsNumberCells()
        {
            var logger = new Mock<ILogger<MewpCoverageReporterService>>();
            var service = new MewpCoverageReporterService(
                logger.Object,
                new SpreadsheetService(),
                new StylesheetService()
            );
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using (var document = SpreadsheetDocument.Create(
                    tempPath,
                    DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
                ))
                {
                    var model = new MewpCoverageReporterModel
                    {
                        TestPlanName = "MEWP",
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "Customer ID", "SR1001" },
                                { "Title (Customer name)", "Req 1001" },
                                { "Responsibility - SAPWBS (ESUK/IL)", "ESUK" },
                                { "Test case id", 101 },
                                { "Test case title", "TC 101" },
                                { "Number of passed steps", 3 },
                                { "Number of failed steps", 1 },
                                { "Number of not run tests", 0 },
                            },
                            new Dictionary<string, object>
                            {
                                { "Customer ID", "SR1002" },
                                { "Title (Customer name)", "Uncovered" },
                                { "Responsibility - SAPWBS (ESUK/IL)", "IL" },
                                { "Test case id", "" },
                                { "Test case title", "" },
                                { "Number of passed steps", 0 },
                                { "Number of failed steps", 0 },
                                { "Number of not run tests", 0 },
                            },
                        }
                    };

                    service.Insert(document, "MEWP Coverage", model);

                    var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var rows = sheetData.Elements<Row>().ToList();

                    var firstDataRowCells = rows[1].Elements<Cell>().ToList();
                    Assert.Equal(CellValues.Number, firstDataRowCells[3].DataType?.Value);
                    Assert.Equal(CellValues.Number, firstDataRowCells[5].DataType?.Value);
                    Assert.Equal(CellValues.Number, firstDataRowCells[6].DataType?.Value);
                    Assert.Equal(CellValues.Number, firstDataRowCells[7].DataType?.Value);

                    var secondDataRowCells = rows[2].Elements<Cell>().ToList();
                    Assert.Equal(CellValues.String, secondDataRowCells[3].DataType?.Value);
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
    }
}
