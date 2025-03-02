using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Moq;
using System.Reflection;

namespace JsonToWord.Services.Tests
{
    public class HtmlServiceTests : IDisposable
    {
        private readonly HtmlService _sut;
        private readonly Mock<IContentControlService> _contentControlServiceMock;
        private readonly Mock<IDocumentValidatorService> _documentValidatorMock;
        private readonly Mock<ILogger<HtmlService>> _loggerMock;
        private WordprocessingDocument _document;
        private MemoryStream _stream;

        public HtmlServiceTests()
        {
            _contentControlServiceMock = new Mock<IContentControlService>();
            _documentValidatorMock = new Mock<IDocumentValidatorService>();
            _loggerMock = new Mock<ILogger<HtmlService>>();

            _sut = new HtmlService(
                _contentControlServiceMock.Object,
                _documentValidatorMock.Object,
                _loggerMock.Object);

            // Create a test document for use in tests
            _stream = new MemoryStream();
            _document = WordprocessingDocument.Create(_stream, WordprocessingDocumentType.Document);
            var mainPart = _document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
        }

        public void Dispose()
        {
            _document?.Dispose();
            _stream?.Dispose();
        }

        // Helper method to invoke private WrapHtmlWithStyle method via reflection
        private string? InvokeWrapHtmlWithStyle(string originalHtml, string font, uint fontSize)
        {
            var methodInfo = typeof(HtmlService).GetMethod("WrapHtmlWithStyle",
                BindingFlags.NonPublic | BindingFlags.Instance);

            return methodInfo?.Invoke(_sut, new object[] { originalHtml, font, fontSize }) as string;
        }

        // Helper method to invoke private ConvertHtmlToOpenXmlElements method via reflection
        private IEnumerable<OpenXmlCompositeElement>? InvokeConvertHtmlToOpenXmlElements(string html, string font = "Arial", uint fontSize = 12)
        {
            var methodInfo = typeof(HtmlService).GetMethod("ConvertHtmlToOpenXmlElements", 
                BindingFlags.Public | BindingFlags.Instance);

            var wordHtml = new WordHtml
            {
                Html = html,
                Font = font,
                FontSize = fontSize
            };

            return methodInfo?.Invoke(_sut, new object[] { wordHtml, _document }) as IEnumerable<OpenXmlCompositeElement>;
        }

        [Fact]
        public void WrapHtmlWithStyle_WithExistingHtmlTags_AppliesStyleToBodyTag()
        {
            // Arrange
            string originalHtml = "<html><head></head><body>Test content</body></html>";
            string font = "Arial";
            uint fontSize = 12;

            // Act
            string? result = InvokeWrapHtmlWithStyle(originalHtml, font, fontSize);

            // Assert
            Assert.Contains("<body style=\"font-family: Arial, sans-serif; font-size: 12pt;\">", result);
            Assert.Contains("Test content", result);
            Assert.DoesNotContain("<body><body", result); // Ensure no duplicate body tags
        }

        [Fact]
        public void WrapHtmlWithStyle_WithExistingStyleInBodyTag_MergesStyles()
        {
            // Arrange
            string originalHtml = "<html><head></head><body style=\"color: red;\">Test content</body></html>";
            string font = "Calibri";
            uint fontSize = 10;

            // Act
            string? result = InvokeWrapHtmlWithStyle(originalHtml, font, fontSize);

            // Assert
            Assert.Contains("<body style=\"color: red; font-family: Calibri, sans-serif; font-size: 10pt;\">", result);
            Assert.Contains("Test content", result);
        }

        [Fact]
        public void WrapHtmlWithStyle_WithoutHtmlTags_WrapsContentAndAppliesStyles()
        {
            // Arrange
            string originalHtml = "<p>Test paragraph</p><div>Test div</div>";
            string font = "Times New Roman";
            uint fontSize = 14;

            // Act
            string? result = InvokeWrapHtmlWithStyle(originalHtml, font, fontSize);

            // Assert
            Assert.Contains("<html>", result);
            Assert.Contains("</html>", result);
            Assert.Contains("<body style='font-family: Times New Roman, sans-serif; font-size: 14pt;'>", result);
            Assert.Contains("<p style='font-family: Times New Roman, sans-serif; font-size: 14pt;'>Test paragraph</p>", result);
            Assert.Contains("<div style='font-family: Times New Roman, sans-serif; font-size: 14pt;'>Test div</div>", result);
        }

        [Fact]
        public void WrapHtmlWithStyle_WithEmptyHtml_ReturnsWrappedEmptyContent()
        {
            // Arrange
            string originalHtml = "";
            string font = "Arial";
            uint fontSize = 12;

            // Act
            string? result = InvokeWrapHtmlWithStyle(originalHtml, font, fontSize);

            // Assert
            Assert.Contains("<html>", result);
            Assert.Contains("<body style='font-family: Arial, sans-serif; font-size: 12pt;'>", result);
            Assert.Contains("</body>", result);
            Assert.Contains("</html>", result);
        }

        [Theory]
        [InlineData("Calibri", 10)]
        [InlineData("Arial", 12)]
        [InlineData("Times New Roman", 14)]
        public void WrapHtmlWithStyle_WithDifferentFontsAndSizes_AppliesCorrectStyles(string font, uint fontSize)
        {
            // Arrange
            string originalHtml = "<p>Test content</p>";

            // Act
            string? result = InvokeWrapHtmlWithStyle(originalHtml, font, fontSize);

            // Assert
            Assert.Contains($"font-family: {font}, sans-serif; font-size: {fontSize}pt;", result);
        }

        [Fact]
        public void WrapHtmlWithStyle_WithWhitespaceAroundHtmlTags_StillRecognizesHtmlStructure()
        {
            // Arrange
            string originalHtml = "  \n  <html>  \n  <body>Test content</body></html>  \n  ";
            string font = "Arial";
            uint fontSize = 12;

            // Act
            string? result = InvokeWrapHtmlWithStyle(originalHtml, font, fontSize);

            // Assert
            Assert.Contains("<body style=\"font-family: Arial, sans-serif; font-size: 12pt;\">", result);
            Assert.Contains("Test content", result);
        }

        [Fact]
        public void ConvertHtmlToOpenXmlElements_WithNestedTableStructure_ReturnsValidNestedTable()
        {
            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Office2016);
            // Arrange
            string invalidHtml = "<table><tbody><tr><td>Cell 1</td><td>Cell 2</td></tr><tr><td>Cell 3</td><td><table><tr><td>Some image<img src='data:image/png;base64,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'/></td></tr></table></td></tr></tbody></table>";

            // Act
            var result = InvokeConvertHtmlToOpenXmlElements(invalidHtml);

            // Assert
            Assert.NotNull(result);
            var resultList = result?.ToList();
            Assert.True(resultList?.Count > 0);

            if (resultList?.Count > 0)
            {
                foreach (var item in resultList)
                {
                    validator.Validate(item);
                    var errors = validator.Validate(item).ToList();

                    Assert.True(!errors.Any());

                    if(item is Table table)
                    {
                        foreach (var row in table.Elements<TableRow>())
                        {
                            foreach (var cell in row.Elements<TableCell>())
                            {
                                OpenXmlElement? lastChild = cell.LastChild;
                                Assert.True(lastChild is Paragraph);
                            }
                        }
                    }
                }
            }


        }

        [Fact]
        public void ConvertHtmlToOpenXmlElement_DivAndParagraphWithNoContent_EmptyListResult()
        {
            string html = "<div><p></p></div>";

            var result = InvokeConvertHtmlToOpenXmlElements(html);
            Assert.NotNull(result);
            var resultList = result?.ToList();
            Assert.True(resultList?.Count == 0);

        }


        [Fact]
        public void ConvertHtmlToOpenXmlElements_WithNestedListsInvalidStructure_ReturnsErrorHtml()
        {
            // Arrange
            string invalidHtml = "<ul><ol><div>Invalid nesting</div></ol></ul>";

            // Act
            var result = InvokeConvertHtmlToOpenXmlElements(invalidHtml);

            // Assert
            Assert.NotNull(result);
            var resultList = result?.ToList();
            Assert.True(resultList?.Count > 0);
            
            // Verify the error was logged
            _loggerMock.Verify(
                x => x.Log(
                    LogLevel.Error,
                    It.IsAny<EventId>(),
                    It.Is<It.IsAnyType>((o, t) => o.ToString().Contains("DocGen ran into an issue parsing the html")),
                    It.IsAny<Exception>(),
                    It.IsAny<Func<It.IsAnyType, Exception?, string>>()),
                Times.Once);
        }
    }
}
