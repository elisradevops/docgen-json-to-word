using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Services.ExcelServices;

namespace JsonToWord.Services.Tests
{
    public class SpreadsheetServiceTests
    {
        [Fact]
        public void GetColumnLetter_HandlesMultiLetterColumns()
        {
            var service = new SpreadsheetService();

            Assert.Equal("A", service.GetColumnLetter(1));
            Assert.Equal("Z", service.GetColumnLetter(26));
            Assert.Equal("AA", service.GetColumnLetter(27));
        }

        [Fact]
        public void GetOrCreateWorksheetPart_CreatesNewWithSafeName()
        {
            var service = new SpreadsheetService();

            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            workbookPart.Workbook.AppendChild(new Sheets());

            var worksheetPart = service.GetOrCreateWorksheetPart(workbookPart, "Invalid/Name*[]ThatIsWayTooLongForExcel");

            Assert.NotNull(worksheetPart.Worksheet);
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
            Assert.True(sheet.Name.Value.Length <= 31);
            Assert.DoesNotContain("/", sheet.Name.Value);
            Assert.DoesNotContain("*", sheet.Name.Value);
            Assert.DoesNotContain("[", sheet.Name.Value);
            Assert.DoesNotContain("]", sheet.Name.Value);
        }

        [Fact]
        public void GetOrCreateWorksheetPart_ReturnsExistingSheet()
        {
            var service = new SpreadsheetService();

            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheetPart = workbookPart.AddNewPart<WorksheetPart>();
            sheetPart.Worksheet = new Worksheet(new SheetData());
            workbookPart.Workbook.AppendChild(new Sheets());
            workbookPart.Workbook.Sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(sheetPart),
                SheetId = 1,
                Name = "Sheet1"
            });

            var result = service.GetOrCreateWorksheetPart(workbookPart, "Sheet1");

            Assert.Same(sheetPart, result);
        }

        [Fact]
        public void CreateHyperlinkCell_AddsHyperlinkWithMappedStyle()
        {
            var service = new SpreadsheetService();

            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var cell = service.CreateHyperlinkCell(
                worksheetPart,
                "A1",
                "Link",
                "https://example.com",
                6,
                "Open link");

            Assert.Equal(12u, cell.StyleIndex.Value);
            var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
            Assert.NotNull(hyperlinks);
            Assert.Contains(hyperlinks.Elements<Hyperlink>(), h => h.Reference == "A1");
        }

        [Fact]
        public void CreateHeaderRow_CreatesGroupAndFieldHeaders()
        {
            var service = new SpreadsheetService();
            var sheetData = new SheetData();
            var mergeCells = new MergeCells();

            var columns = new List<ColumnDefinition>
            {
                new ColumnDefinition { Name = "Id", Property = "TestCaseId", Group = "Test Cases" },
                new ColumnDefinition { Name = "Name", Property = "TestCaseName", Group = "Test Cases" }
            };

            var columnCounts = new Dictionary<string, int>
            {
                { "Test Cases", 2 }
            };
            var groupItemCounts = new Dictionary<string, int>
            {
                { "Test Cases", 3 }
            };

            service.CreateHeaderRow(sheetData, columns, mergeCells, columnCounts, groupItemCounts);

            var rows = sheetData.Elements<Row>().ToList();
            Assert.Equal(2, rows.Count);
            Assert.Contains(rows[1].Elements<Cell>(), c => c.CellValue?.Text == "Id");
            Assert.Contains(rows[1].Elements<Cell>(), c => c.CellValue?.Text == "Name");
            Assert.True(mergeCells.Elements<MergeCell>().Any());
        }
    }
}
