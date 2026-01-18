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

namespace JsonToWord.Services.Tests
{
    public class VoidListServiceTests
    {
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
    }
}
