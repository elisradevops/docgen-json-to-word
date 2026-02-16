using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services;
using JsonToWord.Services.ExcelServices;
using Microsoft.Extensions.Logging;
using Moq;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace JsonToWord.Services.Tests
{
    public class VoidListServiceTests
    {
        [Fact]
        public void CreateVoidList_ParsesCodeAndValue_WhenValueStartsAfterLineBreak()
        {
            var logger = new Mock<ILogger<VoidListService>>();
            var service = new VoidListService(logger.Object, new SpreadsheetService());

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var docPath = Path.Combine(tempDir, "input.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                            new Run(new Text("Test Case - 100"))
                        )
                    );

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new Run(new Text("#VL-12")),
                            new Run(new Break()),
                            new Run(new Text("Very long expected action value#"))
                        )
                    );
                }

                var files = service.CreateVoidList(docPath);
                var voidListPath = files.First(f => f.EndsWith("VOID LIST.xlsx", StringComparison.OrdinalIgnoreCase));

                using var spreadsheet = SpreadsheetDocument.Open(voidListPath, false);
                var row2 = GetRow(spreadsheet, 2);
                Assert.Equal("VL-12", GetCellStringValue(spreadsheet, row2, "A"));
                Assert.Equal("Very long expected action value", GetCellStringValue(spreadsheet, row2, "B"));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateVoidList_ParsesCodeAndValue_WhenValueStartsAfterNewlineCharacter()
        {
            var logger = new Mock<ILogger<VoidListService>>();
            var service = new VoidListService(logger.Object, new SpreadsheetService());

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var docPath = Path.Combine(tempDir, "input.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                            new Run(new Text("Test Case - 200"))
                        )
                    );

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new Run(
                                new Text("#VL-34\nAnother long value spanning multiple words#")
                                {
                                    Space = SpaceProcessingModeValues.Preserve
                                }
                            )
                        )
                    );
                }

                var files = service.CreateVoidList(docPath);
                var voidListPath = files.First(f => f.EndsWith("VOID LIST.xlsx", StringComparison.OrdinalIgnoreCase));

                using var spreadsheet = SpreadsheetDocument.Open(voidListPath, false);
                var row2 = GetRow(spreadsheet, 2);
                Assert.Equal("VL-34", GetCellStringValue(spreadsheet, row2, "A"));
                Assert.Equal("Another long value spanning multiple words", GetCellStringValue(spreadsheet, row2, "B"));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateVoidList_ParsesCodeAndValue_WhenKeyAndValueSplitAcrossRunsWithoutDelimiter()
        {
            var logger = new Mock<ILogger<VoidListService>>();
            var service = new VoidListService(logger.Object, new SpreadsheetService());

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var docPath = Path.Combine(tempDir, "input.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                            new Run(new Text("Test Case - 201"))
                        )
                    );

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new Run(new Text("example value is ")),
                            new Run(new Text("#VL-19")),
                            new Run(new Text("45#"))
                        )
                    );
                }

                var files = service.CreateVoidList(docPath);
                var voidListPath = files.First(f => f.EndsWith("VOID LIST.xlsx", StringComparison.OrdinalIgnoreCase));

                using var spreadsheet = SpreadsheetDocument.Open(voidListPath, false);
                var row2 = GetRow(spreadsheet, 2);
                Assert.Equal("VL-19", GetCellStringValue(spreadsheet, row2, "A"));
                Assert.Equal("45", GetCellStringValue(spreadsheet, row2, "B"));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateVoidList_ParsesWholeNumberCode_WhenNoRunBoundaryExists()
        {
            var logger = new Mock<ILogger<VoidListService>>();
            var service = new VoidListService(logger.Object, new SpreadsheetService());

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var docPath = Path.Combine(tempDir, "input.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                            new Run(new Text("Test Case - 202"))
                        )
                    );

                    mainPart.Document.Body.Append(new Paragraph(new Run(new Text("#VL-1945#"))));
                }

                var files = service.CreateVoidList(docPath);
                var voidListPath = files.First(f => f.EndsWith("VOID LIST.xlsx", StringComparison.OrdinalIgnoreCase));

                using var spreadsheet = SpreadsheetDocument.Open(voidListPath, false);
                var row2 = GetRow(spreadsheet, 2);
                Assert.Equal("VL-1945", GetCellStringValue(spreadsheet, row2, "A"));
                Assert.Equal(string.Empty, GetCellStringValue(spreadsheet, row2, "B"));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateVoidList_CreatesSpreadsheetAndValidationReport()
        {
            var logger = new Mock<ILogger<VoidListService>>();
            var service = new VoidListService(logger.Object, new SpreadsheetService());

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var docPath = Path.Combine(tempDir, "input.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    var heading = new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                        new Run(new Text("Test Case - 100"))
                    );
                    mainPart.Document.Body.Append(heading);

                    mainPart.Document.Body.Append(new Paragraph(new Run(new Text("#VL-1 First value#"))));
                    mainPart.Document.Body.Append(new Paragraph(new Run(new Text("#VL-1 Second value#"))));
                    mainPart.Document.Body.Append(new Paragraph(new Run(new Text("#VL-ABC#"))));
                }

                var files = service.CreateVoidList(docPath);

                Assert.True(files.Any(f => f.EndsWith("VOID LIST.xlsx", StringComparison.OrdinalIgnoreCase)));
                Assert.True(files.Any(f => f.EndsWith("VALIDATION REPORT.txt", StringComparison.OrdinalIgnoreCase)));
                Assert.All(files, f => Assert.True(File.Exists(f)));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateVoidList_SupportsDecimalCode()
        {
            var logger = new Mock<ILogger<VoidListService>>();
            var service = new VoidListService(logger.Object, new SpreadsheetService());

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var docPath = Path.Combine(tempDir, "input.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                            new Run(new Text("Test Case - 300"))
                        )
                    );

                    mainPart.Document.Body.Append(new Paragraph(new Run(new Text("#VL-12.5 Decimal key#"))));
                }

                var files = service.CreateVoidList(docPath);

                var voidListPath = files.First(f => f.EndsWith("VOID LIST.xlsx", StringComparison.OrdinalIgnoreCase));
                using var spreadsheet = SpreadsheetDocument.Open(voidListPath, false);
                var row2 = GetRow(spreadsheet, 2);
                Assert.Equal("VL-12.5", GetCellStringValue(spreadsheet, row2, "A"));
                Assert.Equal("Decimal key", GetCellStringValue(spreadsheet, row2, "B"));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateVoidList_TreatsMalformedDecimalCodeAsInvalid()
        {
            var logger = new Mock<ILogger<VoidListService>>();
            var service = new VoidListService(logger.Object, new SpreadsheetService());

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var docPath = Path.Combine(tempDir, "input.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(docPath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                            new Run(new Text("Test Case - 301"))
                        )
                    );

                    mainPart.Document.Body.Append(new Paragraph(new Run(new Text("#VL-12.5.7 Bad decimal key#"))));
                }

                var files = service.CreateVoidList(docPath);

                Assert.DoesNotContain(files, f => f.EndsWith("VOID LIST.xlsx", StringComparison.OrdinalIgnoreCase));
                var validationReport = files.First(f =>
                    f.EndsWith("VALIDATION REPORT.txt", StringComparison.OrdinalIgnoreCase)
                );
                var reportText = File.ReadAllText(validationReport);
                Assert.Contains("Invalid VL code found: #VL-12.5.7 Bad decimal key#", reportText);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        private static Spreadsheet.Row GetRow(SpreadsheetDocument spreadsheet, uint rowIndex)
        {
            var worksheetPart = spreadsheet.WorkbookPart?.WorksheetParts.First();
            Assert.NotNull(worksheetPart);

            var sheetData = worksheetPart!.Worksheet.GetFirstChild<Spreadsheet.SheetData>();
            Assert.NotNull(sheetData);

            var row = sheetData!.Elements<Spreadsheet.Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            Assert.NotNull(row);
            return row!;
        }

        private static string GetCellStringValue(SpreadsheetDocument spreadsheet, Spreadsheet.Row row, string columnLetter)
        {
            var cell = row.Elements<Spreadsheet.Cell>().FirstOrDefault(c => c.CellReference?.Value?.StartsWith(columnLetter, StringComparison.OrdinalIgnoreCase) == true);
            Assert.NotNull(cell);

            if (cell!.DataType?.Value == Spreadsheet.CellValues.SharedString)
            {
                var sst = spreadsheet.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                Assert.NotNull(sst);
                int sharedIndex = int.Parse(cell.CellValue?.Text ?? "0");
                return sst!.Elements<Spreadsheet.SharedStringItem>().ElementAt(sharedIndex).InnerText;
            }

            return cell.CellValue?.Text ?? string.Empty;
        }
    }
}
