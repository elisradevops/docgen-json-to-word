using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Moq;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace JsonToWord.Services.Tests
{
    public class PictureServiceTests : IDisposable
    {
        private readonly Mock<IContentControlService> _mockContentControlService;
        private readonly Mock<IParagraphService> _mockParagraphService;
        private readonly PictureService _pictureService;

        private readonly string _docPath;
        private readonly string _testImagePath;
        private WordprocessingDocument _document;

        public PictureServiceTests()
        {
            _mockContentControlService = new Mock<IContentControlService>();
            _mockParagraphService = new Mock<IParagraphService>();

            _pictureService = new PictureService(
                _mockContentControlService.Object,
                _mockParagraphService.Object
            );

            // Create a temporary document for testing
            _docPath = Path.Combine(Path.GetTempPath(), $"test_doc_{Guid.NewGuid()}.docx");
            using (var fs = File.Create(_docPath))
            {
                _document = WordprocessingDocument.Create(fs, WordprocessingDocumentType.Document);
                var mainPart = _document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                _document.Save();
            }

            _document = WordprocessingDocument.Open(_docPath, true);

            // Create a real test image file using ImageSharp
            _testImagePath = Path.Combine(Path.GetTempPath(), $"test_image_{Guid.NewGuid()}.png");
            CreateValidTestImage(_testImagePath);

            // Setup common mocks
            _mockContentControlService.Setup(m => m.FindContentControl(_document, It.IsAny<string>()))
                .Returns(new SdtBlock(new SdtProperties(), new SdtEndCharProperties(), new SdtContentBlock()));

            _mockParagraphService.Setup(m => m.CreateCaption(It.IsAny<string>()))
                .Returns(new Paragraph(new Run(new Text("Test Caption"))));
        }

        [Fact]
        public void Insert_WithValidImage_InsertsImageIntoContentControl()
        {
            // Arrange
            var sdtBlock = new SdtBlock(new SdtProperties(), new SdtEndCharProperties());

            _mockContentControlService.Setup(m => m.FindContentControl(_document, "TestControl"))
                .Returns(sdtBlock);

            var wordAttachment = new WordAttachment
            {
                Path = _testImagePath,
                Name = "Test Image"
            };

            // Act
            _pictureService.Insert(_document, "TestControl", wordAttachment);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);
            _mockParagraphService.Verify(m => m.CreateCaption("Test Image"), Times.Once);

            // Get the content block that was ADDED to the sdtBlock
            var addedContentBlock = sdtBlock.Elements<SdtContentBlock>().FirstOrDefault();
            Assert.NotNull(addedContentBlock);

            // Verify content contains paragraph with drawing and caption
            var paragraphs = addedContentBlock?.Elements<Paragraph>().ToList();
            Assert.Equal(2, paragraphs?.Count);

            // First paragraph should contain drawing
            var drawing = paragraphs?[0].Descendants<Drawing>().FirstOrDefault();
            Assert.NotNull(drawing);

            // Second paragraph should be caption
            var captionParagraph = paragraphs?[1];
            Assert.NotNull(captionParagraph);
        }

        [Fact]
        public void CreateDrawing_WithNormalImage_CreatesProperDrawing()
        {
            // Act
            var drawing = _pictureService.CreateDrawing(_document.MainDocumentPart, _testImagePath);

            // Assert
            Assert.NotNull(drawing);

            var inline = drawing.Descendants<DW.Inline>().FirstOrDefault();
            Assert.NotNull(inline);

            // Verify drawing properties
            var docProperties = inline?.Descendants<DW.DocProperties>().FirstOrDefault();
            Assert.NotNull(docProperties);
            Assert.Equal(1U, docProperties?.Id?.Value);  // First image ID = 1

            // Verify blip reference
            var blip = inline?.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            Assert.NotNull(blip);
            Assert.NotNull(blip?.Embed);

            // Verify image part was added
            var imageParts = _document.MainDocumentPart?.ImageParts.ToList();
            Assert.Single(imageParts);
        }

        [Fact]
        public void CreateDrawing_WithFlattenedImage_ReducesImageSize()
        {
            // Act
            var normalDrawing = _pictureService.CreateDrawing(_document.MainDocumentPart, _testImagePath, false);
            var flattenedDrawing = _pictureService.CreateDrawing(_document.MainDocumentPart, _testImagePath, true);

            // Assert
            var normalExtent = normalDrawing.Descendants<DW.Extent>().FirstOrDefault();
            var flattenedExtent = flattenedDrawing.Descendants<DW.Extent>().FirstOrDefault();

            Assert.NotNull(normalExtent);
            Assert.NotNull(flattenedExtent);

            // Flattened drawing should have half the dimensions
            Assert.True(flattenedExtent?.Cx?.Value < normalExtent?.Cx?.Value);
            Assert.True(flattenedExtent?.Cy?.Value < normalExtent?.Cy?.Value);
        }

        [Fact]
        public void CreateDrawing_MultipleCalls_IncrementsImageIds()
        {
            // Act
            var drawing1 = _pictureService.CreateDrawing(_document.MainDocumentPart, _testImagePath);
            var drawing2 = _pictureService.CreateDrawing(_document.MainDocumentPart, _testImagePath);
            var drawing3 = _pictureService.CreateDrawing(_document.MainDocumentPart, _testImagePath);

            // Assert
            var docProps1 = drawing1.Descendants<DW.DocProperties>().FirstOrDefault();
            var docProps2 = drawing2.Descendants<DW.DocProperties>().FirstOrDefault();
            var docProps3 = drawing3.Descendants<DW.DocProperties>().FirstOrDefault();

            Assert.NotNull(docProps1);
            Assert.NotNull(docProps2);
            Assert.NotNull(docProps3);

            // Verify IDs are incrementing
            Assert.True(docProps2?.Id?.Value > docProps1?.Id?.Value);
            Assert.True(docProps3?.Id?.Value > docProps2?.Id?.Value);
        }

        [Fact]
        public void CreateDrawing_WithExistingImages_ContinuesIdSequence()
        {
            // Arrange - Add a DocProperties element with ID=5 to the document
            var paragraph = new Paragraph(
                new Run(
                    new Drawing(
                        new DW.Inline(
                            new DW.DocProperties { Id = 5U }
                        )
                    )
                )
            );
            _document.MainDocumentPart?.Document.Body?.AppendChild(paragraph);

            // Initialize _currentId by first inserting an image through Insert() method
            var initialAttachment = new WordAttachment
            {
                Path = _testImagePath,
                Name = "Initial Image"
            };

            var mockSdtBlock = new SdtBlock();
            _mockContentControlService.Setup(m => m.FindContentControl(_document, "TestControl"))
                .Returns(mockSdtBlock);

            // This will call GetMaxImageId and set _currentId properly
            _pictureService.Insert(_document, "TestControl", initialAttachment);

            // Reset the mock to avoid issues with multiple calls
            _mockContentControlService.Setup(m => m.FindContentControl(_document, "TestControl"))
                .Returns(new SdtBlock());

            // Act - Now call CreateDrawing again to get another drawing
            var drawing = _pictureService.CreateDrawing(_document.MainDocumentPart, _testImagePath);

            // Assert
            var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
            Assert.NotNull(docProps);

            // ID should be greater than 5 and also greater than the ID used in the first image
            Assert.True(docProps?.Id?.Value > 5U);
        }

        [Fact]
        public void Insert_WithNonexistentPath_ThrowsException()
        {
            // Arrange
            var wordAttachment = new WordAttachment
            {
                Path = "NonExistentPath.jpg",
                Name = "Test Image"
            };

            // Act & Assert
            var exception = Assert.ThrowsAny<Exception>(() =>
                _pictureService.Insert(_document, "TestControl", wordAttachment));

            // Verify it's a file not found type of exception
            Assert.True(exception is FileNotFoundException ||
                       exception.Message.Contains("not exist") ||
                       exception.Message.Contains("not found"));
        }

        [Fact]
        public void Insert_WithContentControlNotFound_ThrowsException()
        {
            // Arrange
            _mockContentControlService.Setup(m => m.FindContentControl(_document, "NonExistentControl"))
            .Returns((SdtBlock?)null!);

            var wordAttachment = new WordAttachment
            {
                Path = _testImagePath,
                Name = "Test Image"
            };

            // Act & Assert
            Assert.ThrowsAny<Exception>(() =>
                _pictureService.Insert(_document, "NonExistentControl", wordAttachment));
        }

        /// <summary>
        /// Creates a valid test image using ImageSharp
        /// </summary>
        private void CreateValidTestImage(string path)
        {
            // Create a simple 100x100 image
            using (var image = new Image<Rgba32>(100, 100))
            {
                // Fill with a color
                for (int y = 0; y < image.Height; y++)
                {
                    for (int x = 0; x < image.Width; x++)
                    {
                        image[x, y] = new Rgba32(200, 150, 100);
                    }
                }

                // Save as PNG
                image.Save(path);
            }
        }

        public void Dispose()
        {
            _document?.Dispose();
            if (File.Exists(_docPath))
            {
                File.Delete(_docPath);
            }

            if (File.Exists(_testImagePath))
            {
                File.Delete(_testImagePath);
            }
        }
    }
}