using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Services.ExcelServices;

namespace JsonToWord.Services.Tests
{
    public class StylesheetServiceTests
    {
        [Fact]
        public void CreateStylesheet_DefinesExpectedCounts()
        {
            var service = new StylesheetService();
            var stylesheet = service.CreateStylesheet();

            Assert.NotNull(stylesheet.Fonts);
            Assert.NotNull(stylesheet.Fills);
            Assert.NotNull(stylesheet.CellFormats);
            Assert.Equal(4, stylesheet.Fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().Count());
            Assert.Equal(10, stylesheet.Fills.Elements<DocumentFormat.OpenXml.Spreadsheet.Fill>().Count());
            Assert.Equal(18, stylesheet.CellFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().Count());
        }

        [Fact]
        public void EnsureStylesheet_AddsStylesheetOnce()
        {
            var service = new StylesheetService();

            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            service.EnsureStylesheet(workbookPart);
            service.EnsureStylesheet(workbookPart);

            Assert.Single(workbookPart.GetPartsOfType<WorkbookStylesPart>());
        }
    }
}
