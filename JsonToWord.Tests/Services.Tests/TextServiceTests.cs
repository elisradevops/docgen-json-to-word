using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
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
    }
}
