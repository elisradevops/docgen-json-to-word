using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services;

namespace JsonToWord.Services.Tests
{
    public class UtilsServiceTests
    {
        [Fact]
        public void ConvertCmToDxa_ConvertsWithRounding()
        {
            var service = new UtilsService();
            var dxa = service.ConvertCmToDxa(1);

            Assert.Equal(567, dxa);
        }

        [Fact]
        public void ConvertDxaToPct_ConvertsUsingPageWidth()
        {
            var service = new UtilsService();
            var pct = service.ConvertDxaToPct(2500, 5000);

            Assert.Equal(2500, pct);
        }

        [Fact]
        public void GetPageWidthDxa_ReturnsDefaultWhenMissing()
        {
            var service = new UtilsService();

            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));

            var width = service.GetPageWidthDxa(mainPart);

            Assert.Equal(11906, width);
        }

        [Fact]
        public void GetPageWidthDxa_ReturnsSectionWidth()
        {
            var service = new UtilsService();

            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            var sectionProps = new SectionProperties(new PageSize { Width = 10000u });
            mainPart.Document = new Document(new Body(new Paragraph(), sectionProps));

            var width = service.GetPageWidthDxa(mainPart);

            Assert.Equal(10000, width);
        }

        [Fact]
        public void ParseStringToDouble_ParsesNumericSubstring()
        {
            var service = new UtilsService();
            var value = service.ParseStringToDouble("Width: 12.5cm");

            Assert.Equal(12.5, value);
        }

        [Fact]
        public void ParseStringToDouble_ThrowsForInvalidInput()
        {
            var service = new UtilsService();
            Assert.Throws<FormatException>(() => service.ParseStringToDouble("no numbers here"));
        }
    }
}
