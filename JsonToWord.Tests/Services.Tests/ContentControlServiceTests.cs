using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class ContentControlServiceTests
    {
        [Fact]
        public void FindContentControl_ReturnsExistingBlock()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "cc1" }, new Tag { Val = "cc1" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("content"))))
            );
            mainPart.Document.Body.Append(sdtBlock);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            var result = service.FindContentControl(document, "cc1");

            Assert.Same(sdtBlock, result);
        }

        [Fact]
        public void FindContentControl_ConvertsRunToBlock()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtRun = new SdtRun(
                new SdtProperties(new SdtAlias { Val = "cc2" }, new Tag { Val = "cc2" }),
                new SdtContentRun(new Run(new Text("placeholder")))
            );
            var paragraph = new Paragraph(new Run(new Text("Label ")), sdtRun);
            mainPart.Document.Body.Append(paragraph);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            var result = service.FindContentControl(document, "cc2");

            Assert.NotNull(result);
            Assert.False(mainPart.Document.Body.Descendants<SdtRun>().Any());
            Assert.True(mainPart.Document.Body.Elements<SdtBlock>().Any());
            Assert.True(result.Elements<SdtContentBlock>().Any());
        }

        [Fact]
        public void ClearContentControl_RemovesDefaultContentBlock()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "cc3" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("Click or tap here to enter text."))))
            );
            mainPart.Document.Body.Append(sdtBlock);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            service.ClearContentControl(document, "cc3", false);

            Assert.False(sdtBlock.Elements<SdtContentBlock>().Any());
        }

        [Fact]
        public void ClearContentControl_AllowsRunLevelControlWithoutThrowing()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtRun = new SdtRun(
                new SdtProperties(new SdtAlias { Val = "cc-run-clear" }, new Tag { Val = "cc-run-clear" }),
                new SdtContentRun(new Run(new Text("Click or tap here to enter text.")))
            );
            var paragraph = new Paragraph(new Run(new Text("Label ")), sdtRun);
            mainPart.Document.Body.Append(paragraph);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            service.ClearContentControl(document, "cc-run-clear", false);

            Assert.True(mainPart.Document.Body.Descendants<SdtRun>().Any());
        }

        [Fact]
        public void WritePlainTextToContentControl_WritesIntoSdtCell()
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
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }),
                            new Run(new RunProperties(new Bold()), new Text("Click or tap here to enter text."))
                        )
                    )
                )
            );
            var row = new TableRow(sdtCell);
            var table = new Table(row);
            mainPart.Document.Body.Append(table);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            var result = service.WritePlainTextToContentControl(document, "release-file-content-control", "artifact-1.2.3.zip");

            Assert.True(result);
            Assert.Contains("artifact-1.2.3.zip", mainPart.Document.Body.InnerText);
            var resultingParagraph = mainPart.Document.Body.Descendants<TableCell>().First().Elements<Paragraph>().FirstOrDefault();
            Assert.NotNull(resultingParagraph);
            Assert.Equal("Heading2", resultingParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            Assert.True(resultingParagraph.Descendants<RunProperties>().Any(rp => rp.Bold != null));
        }

        [Fact]
        public void WritePlainTextToContentControl_PreservesBlockParagraphFormatting()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "cc-format" }, new Tag { Val = "cc-format" }),
                new SdtContentBlock(
                    new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = "Heading3" }),
                        new Run(new RunProperties(new Italic()), new Text("Click or tap here to enter text."))
                    )
                )
            );
            mainPart.Document.Body.Append(sdtBlock);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            var result = service.WritePlainTextToContentControl(document, "cc-format", "formatted-text");

            Assert.True(result);
            Assert.Contains("formatted-text", mainPart.Document.Body.InnerText);
            var resultingParagraph = sdtBlock.Descendants<Paragraph>().FirstOrDefault();
            Assert.NotNull(resultingParagraph);
            Assert.Equal("Heading3", resultingParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            Assert.True(resultingParagraph.Descendants<RunProperties>().Any(rp => rp.Italic != null));
        }

        [Fact]
        public void ClearContentControl_DoesNotRemove_WhenNotForced()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "cc-force" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("Custom text"))))
            );
            mainPart.Document.Body.Append(sdtBlock);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            service.ClearContentControl(document, "cc-force", false);

            Assert.True(sdtBlock.Elements<SdtContentBlock>().Any());
        }

        [Fact]
        public void ClearContentControl_RemovesWhenForced()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "cc-force" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("Custom text"))))
            );
            mainPart.Document.Body.Append(sdtBlock);

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            service.ClearContentControl(document, "cc-force", true);

            Assert.False(sdtBlock.Elements<SdtContentBlock>().Any());
        }

        [Fact]
        public void ClearContentControl_ThrowsWhenMissing()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            var ex = Assert.Throws<Exception>(() => service.ClearContentControl(document, "missing", false));

            Assert.Contains("Did not find a content control", ex.Message);
        }

        [Fact]
        public void RemoveContentControl_MergesIntoPreviousParagraph_ForReleaseRange()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var previousParagraph = new Paragraph(new Run(new Text("Before ")));
            mainPart.Document.Body.Append(previousParagraph);

            var sdtBlock = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "release-range-content-control" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("After"))))
            );
            mainPart.Document.Body.Append(sdtBlock);

            var validator = new Mock<IDocumentValidatorService>();
            validator
                .Setup(v => v.ValidateInnerElementOfContentControl(It.IsAny<string>(), It.IsAny<OpenXmlElement>()))
                .Returns(new List<string>());

            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            service.RemoveContentControl(document, "release-range-content-control");

            Assert.False(mainPart.Document.Body.Elements<SdtBlock>().Any());
            Assert.Contains("Before", previousParagraph.InnerText);
            Assert.Contains("After", previousParagraph.InnerText);
        }

        [Fact]
        public void RemoveContentControl_ThrowsWhenValidationErrors()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "cc-errors" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("Bad"))))
            );
            mainPart.Document.Body.Append(sdtBlock);

            var validator = new Mock<IDocumentValidatorService>();
            validator
                .Setup(v => v.ValidateInnerElementOfContentControl("cc-errors", It.IsAny<OpenXmlElement>()))
                .Returns(new List<string> { "error" });

            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            var ex = Assert.Throws<Exception>(() => service.RemoveContentControl(document, "cc-errors"));

            Assert.Contains("Content control is not valid", ex.Message);
        }

        [Fact]
        public void RemoveContentControl_UnwrapsSdtCell()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtCell = new SdtCell(
                new SdtProperties(new SdtAlias { Val = "release-file-content-control" }, new Tag { Val = "release-file-content-control" }),
                new SdtContentCell(
                    new TableCell(new Paragraph(new Run(new Text("artifact-1.2.3.zip"))))
                )
            );
            var row = new TableRow(sdtCell);
            var table = new Table(row);
            mainPart.Document.Body.Append(table);

            var validator = new Mock<IDocumentValidatorService>();
            validator
                .Setup(v => v.ValidateInnerElementOfContentControl(It.IsAny<string>(), It.IsAny<OpenXmlElement>()))
                .Returns(new List<string>());

            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            service.RemoveContentControl(document, "release-file-content-control");

            Assert.False(mainPart.Document.Body.Descendants<SdtCell>().Any());
            Assert.Contains("artifact-1.2.3.zip", mainPart.Document.Body.InnerText);
        }

        [Fact]
        public void ContentControlHeadingMap_CanSetAndClear()
        {
            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            service.MapContentControlHeading("cc", false);

            Assert.False(service.GetContentControlHeadingStatus("cc"));

            service.ClearContentControlHeadingMap();

            Assert.True(service.GetContentControlHeadingStatus("cc"));
        }

        [Fact]
        public void GetContentControlHeadingStatus_ReturnsTrueForEmptyTitle()
        {
            var validator = new Mock<IDocumentValidatorService>();
            var logger = new Mock<ILogger<ContentControlService>>();
            var service = new ContentControlService(logger.Object, validator.Object);

            Assert.True(service.GetContentControlHeadingStatus(null));
            Assert.True(service.GetContentControlHeadingStatus(string.Empty));
        }

        [Fact]
        public void IsUnderStandardHeading_ReturnsTrue_WhenTitleMissing()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var sdtBlock = new SdtBlock(new SdtProperties(), new SdtContentBlock());
            mainPart.Document.Body.Append(sdtBlock);

            var service = new ContentControlService(new Mock<ILogger<ContentControlService>>().Object, new Mock<IDocumentValidatorService>().Object);

            var result = service.IsUnderStandardHeading(sdtBlock);

            Assert.True(result);
        }

        [Fact]
        public void IsUnderStandardHeading_ReturnsTrue_ForStandardHeading()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
            stylesPart.Styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", CustomStyle = false });

            var heading = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Heading")));
            var sdtBlock = new SdtBlock(new SdtProperties(new SdtAlias { Val = "cc" }), new SdtContentBlock());
            mainPart.Document.Body.Append(heading, sdtBlock);

            var service = new ContentControlService(new Mock<ILogger<ContentControlService>>().Object, new Mock<IDocumentValidatorService>().Object);

            var result = service.IsUnderStandardHeading(sdtBlock);

            Assert.True(result);
        }

        [Fact]
        public void IsUnderStandardHeading_ReturnsFalse_ForCustomHeading()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
            stylesPart.Styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", CustomStyle = true });

            var heading = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Heading")));
            var sdtBlock = new SdtBlock(new SdtProperties(new SdtAlias { Val = "cc" }), new SdtContentBlock());
            mainPart.Document.Body.Append(heading, sdtBlock);

            var service = new ContentControlService(new Mock<ILogger<ContentControlService>>().Object, new Mock<IDocumentValidatorService>().Object);

            var result = service.IsUnderStandardHeading(sdtBlock);

            Assert.False(result);
        }

        [Fact]
        public void IsUnderStandardHeading_UsesMappedStatus()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var first = new SdtBlock(new SdtProperties(new SdtAlias { Val = "first" }), new SdtContentBlock());
            var second = new SdtBlock(new SdtProperties(new SdtAlias { Val = "second" }), new SdtContentBlock());
            mainPart.Document.Body.Append(first, second);
            var secondInDoc = mainPart.Document.Body.Elements<SdtBlock>().Last();

            var service = new ContentControlService(new Mock<ILogger<ContentControlService>>().Object, new Mock<IDocumentValidatorService>().Object);
            service.MapContentControlHeading("second", false);

            var result = service.IsUnderStandardHeading(secondInDoc);

            Assert.False(result);
        }

        [Fact]
        public void IsUnderStandardHeading_UsesPreviousHeadingWhenNotMapped()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
            stylesPart.Styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", CustomStyle = true });

            var heading = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Heading")));
            var first = new SdtBlock(new SdtProperties(new SdtAlias { Val = "first" }), new SdtContentBlock());
            var second = new SdtBlock(new SdtProperties(new SdtAlias { Val = "second" }), new SdtContentBlock());
            mainPart.Document.Body.Append(heading, first, second);

            var service = new ContentControlService(new Mock<ILogger<ContentControlService>>().Object, new Mock<IDocumentValidatorService>().Object);

            var result = service.IsUnderStandardHeading(second);

            Assert.False(result);
        }

        [Fact]
        public void IsUnderStandardHeading_UsesPreviousHeadingWhenNotMapped_ReturnsTrueForStandardHeading()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
            stylesPart.Styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", CustomStyle = false });

            var heading = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), new Run(new Text("Heading")));
            var first = new SdtBlock(new SdtProperties(new SdtAlias { Val = "first" }), new SdtContentBlock());
            var second = new SdtBlock(new SdtProperties(new SdtAlias { Val = "second" }), new SdtContentBlock());
            mainPart.Document.Body.Append(heading, first, second);

            var service = new ContentControlService(new Mock<ILogger<ContentControlService>>().Object, new Mock<IDocumentValidatorService>().Object);

            var result = service.IsUnderStandardHeading(second);

            Assert.True(result);
        }
    }
}
