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
    public class FlatTestReporterServiceTests
    {
        [Fact]
        public void Insert_UsesCustomColumnOrder_WhenProvided()
        {
            var logger = new Mock<ILogger<FlatTestReporterService>>();
            var service = new FlatTestReporterService(logger.Object, new SpreadsheetService());
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

            try
            {
                using (var document = SpreadsheetDocument.Create(tempPath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    var model = new FlatTestReporterModel
                    {
                        TestPlanName = "Sheet A",
                        ColumnOrder = new List<string> { "Col A", "Col B" },
                        Rows = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                { "Col A", "A1" },
                                { "Col B", "B1" },
                            }
                        }
                    };

                    service.Insert(document, "Sheet A", model);

                    var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                    var rows = sheetData.Elements<Row>().ToList();
                    var headerValues = rows[0].Elements<Cell>().Select(c => c.CellValue?.Text ?? string.Empty).ToList();
                    var dataValues = rows[1].Elements<Cell>().Select(c => c.CellValue?.Text ?? string.Empty).ToList();

                    Assert.Equal(new List<string> { "Col A", "Col B" }, headerValues);
                    Assert.Equal(new List<string> { "A1", "B1" }, dataValues);
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
