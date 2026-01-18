using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class DocumentValidatorServiceTests
    {
        [Fact]
        public void ValidateDocument_ReturnsTrueForValidDocument()
        {
            var logger = new Mock<ILogger<DocumentValidatorService>>();
            var service = new DocumentValidatorService(logger.Object);

            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("ok")))));

            var isValid = service.ValidateDocument(document);

            Assert.True(isValid);
        }

        [Fact]
        public void ValidateInnerElementOfContentControl_ReturnsEmptyForNull()
        {
            var logger = new Mock<ILogger<DocumentValidatorService>>();
            var service = new DocumentValidatorService(logger.Object);

            var errors = service.ValidateInnerElementOfContentControl("cc", null);

            Assert.Empty(errors);
        }

        [Fact]
        public void ValidateInnerElementOfContentControl_ReturnsEmptyForAltChunk()
        {
            var logger = new Mock<ILogger<DocumentValidatorService>>();
            var service = new DocumentValidatorService(logger.Object);

            var errors = service.ValidateInnerElementOfContentControl("cc", new AltChunk());

            Assert.Empty(errors);
        }

        [Fact]
        public void ValidateInnerElementOfContentControl_ReturnsNoErrorsForValidElement()
        {
            var logger = new Mock<ILogger<DocumentValidatorService>>();
            var service = new DocumentValidatorService(logger.Object);

            var paragraph = new Paragraph(new Run(new Text("ok")));
            var errors = service.ValidateInnerElementOfContentControl("cc", paragraph);

            Assert.Empty(errors);
        }
    }
}
