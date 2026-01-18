using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class DocumentServiceTests
    {
        [Fact]
        public void CreateDocument_ThrowsForUnsupportedExtension()
        {
            var logger = new Mock<ILogger<DocumentService>>();
            var service = new DocumentService(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var path = Path.Combine(tempDir, "template.txt");
            File.WriteAllText(path, "not a docx");

            try
            {
                Assert.Throws<Exception>(() => service.CreateDocument(path));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateDocument_FromDotx_CreatesDocxCopy()
        {
            var logger = new Mock<ILogger<DocumentService>>();
            var service = new DocumentService(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var templatePath = Path.Combine(tempDir, "template.dotx");

            try
            {
                using (var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Template))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("template")))));
                }

                var resultPath = service.CreateDocument(templatePath);

                Assert.EndsWith(".docx", resultPath, StringComparison.OrdinalIgnoreCase);
                Assert.True(File.Exists(resultPath));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void SetLandscape_AddsSectionPropsAndPageSize_WhenMissing()
        {
            var logger = new Mock<ILogger<DocumentService>>();
            var service = new DocumentService(logger.Object);

            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));

            service.SetLandscape(mainPart);

            var sectionProps = mainPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
            Assert.NotNull(sectionProps);

            var pageSize = sectionProps.GetFirstChild<PageSize>();
            Assert.NotNull(pageSize);
            Assert.Equal(PageOrientationValues.Landscape, pageSize.Orient.Value);
            Assert.Equal(16840u, pageSize.Width.Value);
            Assert.Equal(11906u, pageSize.Height.Value);
        }

        [Fact]
        public void SetLandscape_UpdatesExistingPageSize()
        {
            var logger = new Mock<ILogger<DocumentService>>();
            var service = new DocumentService(logger.Object);

            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            var sectionProps = new SectionProperties();
            sectionProps.Append(new PageSize { Orient = PageOrientationValues.Portrait, Width = 100u, Height = 200u });
            mainPart.Document = new Document(new Body(new Paragraph(), sectionProps));

            service.SetLandscape(mainPart);

            var pageSize = mainPart.Document.Body.Elements<SectionProperties>().First().GetFirstChild<PageSize>();
            Assert.NotNull(pageSize);
            Assert.Equal(PageOrientationValues.Landscape, pageSize.Orient.Value);
            Assert.Equal(16840u, pageSize.Width.Value);
            Assert.Equal(11906u, pageSize.Height.Value);
        }
    }
}
