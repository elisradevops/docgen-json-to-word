using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class TextServiceTests
    {
        [Fact]
        public void Write_WithValidUri_AddsHyperlink()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(new SdtProperties(new SdtAlias { Val = "cc" }), new SdtContentBlock());
            mainPart.Document.Body.Append(sdtBlock);

            var contentControlService = new Mock<IContentControlService>();
            contentControlService
                .Setup(s => s.FindContentControl(document, "cc"))
                .Returns(sdtBlock);

            var paragraphService = new ParagraphService();
            var runService = new RunService(new Mock<IPictureService>().Object);
            var service = new TextService(contentControlService.Object, paragraphService, runService);

            var wordParagraph = new WordParagraph
            {
                Runs = new List<WordRun>
                {
                    new WordRun { Text = "Link", Uri = "https://example.com" }
                }
            };

            service.Write(document, "cc", wordParagraph, true);

            Assert.True(sdtBlock.Descendants<Hyperlink>().Any());
        }

        [Fact]
        public void Write_WithInvalidUri_AddsRunWithoutHyperlink()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(new SdtProperties(new SdtAlias { Val = "cc" }), new SdtContentBlock());
            mainPart.Document.Body.Append(sdtBlock);

            var contentControlService = new Mock<IContentControlService>();
            contentControlService
                .Setup(s => s.FindContentControl(document, "cc"))
                .Returns(sdtBlock);

            var paragraphService = new ParagraphService();
            var runService = new RunService(new Mock<IPictureService>().Object);
            var service = new TextService(contentControlService.Object, paragraphService, runService);

            var wordParagraph = new WordParagraph
            {
                Runs = new List<WordRun>
                {
                    new WordRun { Text = "Bad Link", Uri = "not a uri" }
                }
            };

            service.Write(document, "cc", wordParagraph, true);

            Assert.False(sdtBlock.Descendants<Hyperlink>().Any());
            Assert.True(sdtBlock.Descendants<Run>().Any());
        }

        [Fact]
        public void Write_ToSdtCell_PreservesRunStyling()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtCell = new SdtCell(
                new SdtProperties(new SdtAlias { Val = "release-file-content-control" }, new Tag { Val = "release-file-content-control" }),
                new SdtContentCell(
                    new TableCell(
                        new TableCellProperties(new TableCellWidth { Width = "2400", Type = TableWidthUnitValues.Dxa }),
                        new Paragraph(new Run(new Text("Click or tap here to enter text.")))
                    )
                )
            );
            var row = new TableRow(sdtCell);
            var table = new Table(row);
            mainPart.Document.Body.Append(table);

            var contentControlService = new ContentControlService(
                new Mock<ILogger<ContentControlService>>().Object,
                new Mock<IDocumentValidatorService>().Object
            );
            var paragraphService = new ParagraphService();
            var runService = new RunService(new Mock<IPictureService>().Object);
            var service = new TextService(contentControlService, paragraphService, runService);

            var wordParagraph = new WordParagraph
            {
                Runs = new List<WordRun>
                {
                    new WordRun { Text = "test-release-Release-17.zip", Font = "Arial", Size = 10, Bold = true }
                }
            };

            service.Write(document, "release-file-content-control", wordParagraph, true);

            var targetRun = mainPart.Document.Body.Descendants<TableCell>().First().Descendants<Run>().FirstOrDefault();
            Assert.NotNull(targetRun);
            var runProperties = targetRun.RunProperties;
            Assert.NotNull(runProperties);
            Assert.Equal("Arial", runProperties.RunFonts?.Ascii?.Value);
            Assert.Equal("20", runProperties.FontSize?.Val?.Value);
            Assert.NotNull(runProperties.Bold);
            Assert.Contains("test-release-Release-17.zip", mainPart.Document.Body.InnerText);
        }
    }
}
