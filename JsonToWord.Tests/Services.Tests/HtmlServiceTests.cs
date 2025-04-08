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
                It.Is<It.IsAnyType>((o, t) => o != null && o.ToString()!.Contains("DocGen ran into an issue parsing the html")),
                It.IsAny<Exception>(),
                It.IsAny<Func<It.IsAnyType, Exception?, string>>()),
            Times.Once);
        }
    }
}
