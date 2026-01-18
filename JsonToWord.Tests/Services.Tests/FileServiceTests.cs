using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.EventHandlers;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Moq;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace JsonToWord.Services.Tests
{
    [Collection("NonParallel")]
    public class FileServiceTest : IDisposable
    {
        private readonly Mock<IContentControlService> _mockContentControlService;
        private readonly Mock<ILogger<FileService>> _mockLogger;
        private readonly FileService _fileService;

        private readonly string _originalCwd;
        private readonly string _docPath;
        private readonly string _testFilesPath;
        private readonly string _attachmentsPath;
        private readonly string _iconsPath;
        private WordprocessingDocument _document;

        private bool _nonOfficeAttachmentEventFired;

        public FileServiceTest()
        {
            _mockContentControlService = new Mock<IContentControlService>();
            _mockLogger = new Mock<ILogger<FileService>>();

            _fileService = new FileService(
                _mockContentControlService.Object,
                _mockLogger.Object
            );

            // Subscribe to the event
            _fileService.nonOfficeAttachmentEventHandler += () => _nonOfficeAttachmentEventFired = true;

            // Create temporary directory for test files
            _testFilesPath = Path.Combine(Path.GetTempPath(), $"fileservice_test_{Guid.NewGuid()}");
            Directory.CreateDirectory(_testFilesPath);

            // Create temporary attachments folder
            _attachmentsPath = Path.Combine(_testFilesPath, "attachments");
            Directory.CreateDirectory(_attachmentsPath);

            // Create icons directory and copy test icons
            _iconsPath = Path.Combine(_testFilesPath, "Resources", "Icons");
            Directory.CreateDirectory(_iconsPath);
            CreateTestIcons();

            // Set current directory for relative paths to work
            _originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = _testFilesPath;

            // Create temporary document for testing
            _docPath = Path.Combine(_testFilesPath, $"test_doc_{Guid.NewGuid()}.docx");
            using (var fs = File.Create(_docPath))
            {
                _document = WordprocessingDocument.Create(fs, WordprocessingDocumentType.Document);
                var mainPart = _document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                _document.Save();
            }

            _document = WordprocessingDocument.Open(_docPath, true);

            // Setup common mocks
            var sdtBlock = new SdtBlock();
            _mockContentControlService.Setup(m => m.FindContentControl(_document, It.IsAny<string>()))
                .Returns(sdtBlock);
        }

        [Fact]
        public void Insert_WithNormalAttachment_CallsAttachFileToParagraph()
        {
            // Arrange
            var testFilePath = CreateTestFile(".docx", "Test content");
            var wordAttachment = new WordAttachment
            {
                Path = testFilePath,
                Name = "TestDoc.docx",
                IncludeAttachmentContent = false
            };

            var sdtBlock = new SdtBlock();
            _mockContentControlService.Setup(m => m.FindContentControl(_document, "TestControl"))
                .Returns(sdtBlock);

            // Act
            _fileService.Insert(_document, "TestControl", wordAttachment);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);

            // Verify that content was added to the SdtBlock
            Assert.NotEmpty(sdtBlock.ChildElements);
            var contentBlock = sdtBlock.GetFirstChild<SdtContentBlock>();
            Assert.NotNull(contentBlock);

            // Verify paragraph was created
            var paragraph = contentBlock?.GetFirstChild<Paragraph>();
            Assert.NotNull(paragraph);
        }

        [Fact]
        public void Insert_WithIncludeAttachmentContentTrue_CallsAddDocFileContent()
        {
            // Arrange
            var testDocxPath = CreateTestWordDocument();
            var wordAttachment = new WordAttachment
            {
                Path = testDocxPath,
                Name = "TestDoc.docx",
                IncludeAttachmentContent = true
            };

            var sdtBlock = new SdtBlock();
            _mockContentControlService.Setup(m => m.FindContentControl(_document, "TestControl"))
                .Returns(sdtBlock);

            // Act
            _fileService.Insert(_document, "TestControl", wordAttachment);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);

            // Verify content was added to the SdtBlock
            Assert.NotEmpty(sdtBlock.ChildElements);
            var contentBlock = sdtBlock.GetFirstChild<SdtContentBlock>();
            Assert.NotNull(contentBlock);

            // Verify AltChunk was created
            var altChunk = contentBlock?.GetFirstChild<AltChunk>();
            Assert.NotNull(altChunk);
            Assert.NotNull(altChunk?.Id);
            Assert.True(altChunk?.Id?.Value?.StartsWith("altChunkId"));
        }

        [Fact]
        public void AttachFileToParagraph_WithWordDocument_CreatesEmbeddedOfficeFileParagraph()
        {
            // Arrange
            var testFilePath = CreateTestFile(".docx", "Test content");
            var wordAttachment = new WordAttachment
            {
                Path = testFilePath,
                Name = "TestDoc.docx"
            };

            // Act
            var paragraph = _fileService.AttachFileToParagraph(_document.MainDocumentPart, wordAttachment);

            // Assert
            Assert.NotNull(paragraph);

            // Verify embedded object exists
            var embeddedObject = paragraph.Descendants<EmbeddedObject>().FirstOrDefault();
            Assert.NotNull(embeddedObject);

            // Verify OleObject exists with proper attributes
            var oleObject = embeddedObject?.Descendants<OleObject>().FirstOrDefault();
            Assert.NotNull(oleObject);
            Assert.Equal("Word.Document", oleObject?.ProgId);
            Assert.Equal("embed", oleObject?.Type?.ToString()?.ToLower());

            // Verify run with text is present (filename)
            var textRuns = paragraph.Descendants<Run>()
                .Where(r => r.Descendants<Text>().Any(t => t.Text == wordAttachment.Name))
                .ToList();
            Assert.NotEmpty(textRuns);
        }

        [Fact]
        public void AttachFileToParagraph_WithNonOfficeDocument_CreatesHyperlinkParagraph()
        {
            // Arrange
            var testFilePath = CreateTestFile(".pdf", "PDF content");
            var wordAttachment = new WordAttachment
            {
                Path = testFilePath,
                Name = "TestPdf.pdf",
                IsLinkedFile = true
            };

            // Act
            var paragraph = _fileService.AttachFileToParagraph(_document.MainDocumentPart, wordAttachment);

            // Assert
            Assert.NotNull(paragraph);

            // Verify hyperlink exists
            var hyperlink = paragraph.Descendants<Hyperlink>().FirstOrDefault();
            Assert.NotNull(hyperlink);
            Assert.NotNull(hyperlink?.Id);

            // Verify text content
            var hyperlinkText = hyperlink?.Descendants<Text>().FirstOrDefault();
            Assert.NotNull(hyperlinkText);
            Assert.Equal(wordAttachment.Name, hyperlinkText?.Text);

            // Verify icon drawing exists
            var drawing = paragraph.Descendants<Drawing>().FirstOrDefault();
            Assert.NotNull(drawing);

            // Verify event was fired
            Assert.True(_nonOfficeAttachmentEventFired);
        }

        [Fact]
        public void AttachFileToParagraph_WithNullAttachment_ThrowsException()
        {
            // Act & Assert
            var exception = Assert.Throws<Exception>(() =>
                _fileService.AttachFileToParagraph(_document.MainDocumentPart, null));

            Assert.Equal("Word attachment is not defined", exception.Message);
        }

        [Fact]
        public void CreateIconImageDrawing_ReturnsValidDrawing()
        {
            // Arrange
            string imageId = string.Empty;
            var testFilePath = CreateTestFile(".docx", "Test content");
            var wordAttachment = new WordAttachment
            {
                Path = testFilePath,
                Name = "TestDoc.docx"
            };

            // Access the private method using reflection
            MethodInfo createIconMethod = typeof(FileService).GetMethod("CreateIconImageDrawing",
                            BindingFlags.NonPublic | BindingFlags.Instance) ?? throw new InvalidOperationException("Method not found: CreateIconImageDrawing");

            // Act
            var drawing = (Drawing?)createIconMethod.Invoke(_fileService,
                new object[] { _document.MainDocumentPart!, wordAttachment, imageId });

            // Assert
            Assert.NotNull(drawing);

            // Verify Inline element exists
            var inline = drawing?.Descendants<Inline>().FirstOrDefault();
            Assert.NotNull(inline);

            // Verify extent with dimensions
            var extent = inline?.Descendants<Extent>().FirstOrDefault();
            Assert.NotNull(extent);
            Assert.Equal(32 * 9525, extent?.Cx);
            Assert.Equal(32 * 9525, extent?.Cy);

            // Verify blip exists
            var blip = drawing?.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            Assert.NotNull(blip);
        }

        [Fact]
        public void AddDocFileContent_WithValidWordDoc_AddsAltChunk()
        {
            // Arrange
            var testDocxPath = CreateTestWordDocument();

            var wordAttachment = new WordAttachment
            {
                Path = testDocxPath,
                Name = "TestDoc.docx"
            };

            // Access the private method using reflection
            MethodInfo? addDocMethod = typeof(FileService).GetMethod("AddDocFileContent",
                BindingFlags.NonPublic | BindingFlags.Instance);

            // Act
            var altChunk = (AltChunk)addDocMethod?.Invoke(_fileService,
                        new object[] { _document.MainDocumentPart!, wordAttachment })!;

            // Assert
            Assert.NotNull(altChunk);
            Assert.NotNull(altChunk.Id);
            Assert.True(altChunk?.Id?.Value?.StartsWith("altChunkId"));

            // Verify the AlternativeFormatImportPart was added
            var importPart = _document.MainDocumentPart?.AlternativeFormatImportParts
                .FirstOrDefault(p => p.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            Assert.NotNull(importPart);
        }



        [Fact]
        public void PrefixStyleIds_MakesStyleIdsUniqueButPreservesDefaultStyles()
        {
            // Create a test document with styles
            var testDocWithStylesPath = Path.Combine(_testFilesPath, "docWithStyles.docx");
            using (var document = WordprocessingDocument.Create(testDocWithStylesPath, WordprocessingDocumentType.Document))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Add styles part with custom styles
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var styles = new Styles();

                // Add default style that shouldn't be modified
                styles.AppendChild(new Style
                {
                    StyleId = "Normal",
                    Type = StyleValues.Paragraph,
                    Default = true
                });

                // Add custom style that should be prefixed
                styles.AppendChild(new Style
                {
                    StyleId = "CustomStyle1",
                    Type = StyleValues.Paragraph,
                    BasedOn = new BasedOn { Val = "Normal" }
                });

                // Add another custom style referencing the first one
                styles.AppendChild(new Style
                {
                    StyleId = "CustomStyle2",
                    Type = StyleValues.Paragraph,
                    BasedOn = new BasedOn { Val = "CustomStyle1" }
                });

                stylesPart.Styles = styles;

                // Add paragraph using the custom style
                var paragraph = new Paragraph();
                paragraph.ParagraphProperties = new ParagraphProperties
                {
                    ParagraphStyleId = new ParagraphStyleId { Val = "CustomStyle1" }
                };
                mainPart.Document.Body?.AppendChild(paragraph);

                document.Save();
            }

            // Access the private method using reflection
            MethodInfo? prefixStylesMethod = typeof(FileService).GetMethod("PrefixStyleIds",
                        BindingFlags.NonPublic | BindingFlags.Instance) ?? throw new InvalidOperationException("Method not found: PrefixStyleIds");

            // Open the document for testing
            using (var doc = WordprocessingDocument.Open(testDocWithStylesPath, true))
            {
                // Act
                prefixStylesMethod.Invoke(_fileService, new object[] { doc });

                // Save and reopen to ensure changes are applied
                doc.Save();
            }

            // Examine the modified document
            using (var modifiedDoc = WordprocessingDocument.Open(testDocWithStylesPath, false))
            {
                var stylesPart = modifiedDoc.MainDocumentPart?.StyleDefinitionsPart;
                var styles = stylesPart?.Styles;

                // Default styles should remain unchanged
                var normalStyle = styles?.Elements<Style>().FirstOrDefault(s => s.StyleId == "Normal");
                Assert.NotNull(normalStyle);

                // Custom styles should be prefixed
                var customStyle1 = styles?.Elements<Style>().FirstOrDefault(s => s.StyleId != "Normal" && s.StyleId?.Value?.EndsWith("CustomStyle1") == true);
                Assert.NotNull(customStyle1);
                Assert.StartsWith("s", customStyle1?.StyleId);

                // References should be updated too
                var customStyle2 = styles?.Elements<Style>().FirstOrDefault(s => s.StyleId != "Normal" && s.StyleId?.Value?.EndsWith("CustomStyle2") == true);
                Assert.NotNull(customStyle2);
                Assert.Equal(customStyle1?.StyleId, customStyle2?.BasedOn?.Val);

                // Paragraph references should be updated
                var paragraph = modifiedDoc.MainDocumentPart?.Document.Body?.Elements<Paragraph>().First();
                Assert.Equal(customStyle1?.StyleId, paragraph?.ParagraphProperties?.ParagraphStyleId?.Val);
            }
        }

        [Fact]
        public void GetFileContentType_ReturnsCorrectContentType()
        {
            // Access the private method using reflection
            MethodInfo? getContentTypeMethod = typeof(FileService).GetMethod("GetFileContentType",
                BindingFlags.NonPublic | BindingFlags.Instance);

            // Act & Assert - Test different file types
            Assert.Equal("application/msword",
                getContentTypeMethod?.Invoke(_fileService, new[] { "test.doc" }));

            Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                getContentTypeMethod?.Invoke(_fileService, new[] { "test.docx" }));

            Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                getContentTypeMethod?.Invoke(_fileService, new[] { "test.xlsx" }));

            Assert.Equal("application/vnd.ms-powerpoint",
                getContentTypeMethod?.Invoke(_fileService, new[] { "test.ppt" }));

            Assert.Equal("application/vnd.openxmlformats-officedocument.presentationml.presentation",
                getContentTypeMethod?.Invoke(_fileService, new[] { "test.pptx" }));

            Assert.Equal("application/octet-stream",
                getContentTypeMethod?.Invoke(_fileService, new[] { "test.unknown" }));
        }

        [Fact]
        public void GetProdId_ReturnsCorrectProdId()
        {
            // Access the private method using reflection
            MethodInfo? getProdIdMethod = typeof(FileService).GetMethod("GetProdId",
                BindingFlags.NonPublic | BindingFlags.Instance);

            // Act & Assert - Test different file extensions
            Assert.Equal("Word.Document",
                getProdIdMethod?.Invoke(_fileService, new[] { ".docx" }));

            Assert.Equal("Word.Template",
                getProdIdMethod?.Invoke(_fileService, new[] { ".dotx" }));

            Assert.Equal("Excel.Sheet",
                getProdIdMethod?.Invoke(_fileService, new[] { ".xlsx" }));

            Assert.Equal("PowerPoint.Show",
                getProdIdMethod?.Invoke(_fileService, new[] { ".pptx" }));

            Assert.Equal("PowerPoint.Template",
                getProdIdMethod?.Invoke(_fileService, new[] { ".potx" }));

            Assert.Equal("Package",
                getProdIdMethod?.Invoke(_fileService, new[] { ".unknown" }));
        }

        [Theory]
        [InlineData("file.doc", "word.png")]
        [InlineData("file.dotx", "word.png")]
        [InlineData("file.xlsx", "excel.png")]
        [InlineData("file.pptm", "powerpoint.png")]
        [InlineData("file.pot", "powerpoint.png")]
        [InlineData("file.txt", "txt.png")]
        [InlineData("file.pdf", "pdf.png")]
        [InlineData("file.csv", "csv.png")]
        [InlineData("file.jpeg", "picture.png")]
        [InlineData("file.mp3", "media.png")]
        [InlineData("file.xml", "xml.png")]
        [InlineData("file.7z", "zip.png")]
        [InlineData("file.rar", "rar.png")]
        [InlineData("file.unknown", "default.png")]
        [InlineData("FILE.DOCX", "word.png")]
        [InlineData("file", "default.png")]
        public void GetIconPathForFileType_ReturnsExpectedIcon(string fileName, string expectedIcon)
        {
            MethodInfo? getIconPathMethod = typeof(FileService).GetMethod("GetIconPathForFileType",
                BindingFlags.NonPublic | BindingFlags.Instance);

            var result = (string?)getIconPathMethod?.Invoke(_fileService, new object[] { fileName });

            Assert.Equal(expectedIcon, Path.GetFileName(result));
        }

        [Fact]
        public void CopyAttachment_CreatesUniqueFilename_WhenFileExists()
        {
            // Arrange
            var originalFilePath = CreateTestFile(".txt", "Test content");
            var fileName = Path.GetFileNameWithoutExtension(originalFilePath);
            var wordAttachment = new WordAttachment
            {
                Path = originalFilePath,
                Name = fileName
            };

            // Access the private method using reflection
            MethodInfo? copyAttachmentMethod = typeof(FileService).GetMethod("CopyAttachment",
                BindingFlags.NonPublic | BindingFlags.Instance);

            // Act
            var destination1 = (string?)copyAttachmentMethod?.Invoke(_fileService, new object[] { wordAttachment });

            // Now create a second copy with the same name - should get a unique filename
            var destination2 = (string?)copyAttachmentMethod?.Invoke(_fileService, new object[] { wordAttachment });

            // Assert
            Assert.NotEqual(destination1, destination2);
            Assert.True(File.Exists(destination1));
            Assert.True(File.Exists(destination2));
            Assert.Contains("CopyID", destination2); // Should contain the CopyID marker
        }

        [Fact]
        public void TriggerNonOfficeFile_RaisesEvent()
        {
            // Arrange
            bool eventFired = false;
            _fileService.nonOfficeAttachmentEventHandler += () => eventFired = true;

            // Access the private method using reflection
            MethodInfo? triggerMethod = typeof(FileService).GetMethod("TriggerNonOfficeFile",
                BindingFlags.NonPublic | BindingFlags.Instance);

            // Act
            triggerMethod?.Invoke(_fileService, null);

            // Assert
            Assert.True(eventFired);
        }

        public void Dispose()
        {
            _document?.Dispose();

            var restorePath = Directory.Exists(_originalCwd) ? _originalCwd : AppContext.BaseDirectory;
            Environment.CurrentDirectory = restorePath;

            // Clean up temporary directories
            if (Directory.Exists(_testFilesPath))
            {
                try
                {
                    Directory.Delete(_testFilesPath, true);
                }
                catch
                {
                    // Ignore cleanup failures
                }
            }
        }

        #region Helper Methods

        private string CreateTestFile(string extension, string content)
        {
            var filePath = Path.Combine(_testFilesPath, $"testFile{Guid.NewGuid()}{extension}");
            File.WriteAllText(filePath, content);
            return filePath;
        }

        private string CreateTestWordDocument()
        {
            var filePath = Path.Combine(_testFilesPath, $"testDoc{Guid.NewGuid()}.docx");

            using (var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Add a simple paragraph
                var paragraph = new Paragraph(new Run(new Text("Test document content")));
                mainPart.Document.Body?.AppendChild(paragraph);

                document.Save();
            }

            return filePath;
        }

        private void CreateTestIcons()
        {
            // Create mock icons for testing
            var iconTypes = new[] { "word", "excel", "powerpoint", "pdf", "txt", "csv", "picture", "media", "xml", "zip", "rar", "default" };

            foreach (var iconType in iconTypes)
            {
                var iconPath = Path.Combine(_iconsPath, $"{iconType}.png");
                CreateMinimalImage(iconPath);
            }
        }

        private void CreateMinimalImage(string path)
        {
            // Create a minimal valid PNG file for testing
            // Ensure the directory path is not null
            var directoryPath = Path.GetDirectoryName(path);
            if (directoryPath != null)
            {
                Directory.CreateDirectory(directoryPath);
            }


            // PNG header and minimal IHDR chunk for a 1x1 pixel
            byte[] pngData = {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,  // PNG signature
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,  // IHDR chunk length and type
                0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,  // Width=1, Height=1
                0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,  // Bit depth, color type, etc.
                0xDE,                                            // CRC checksum
                0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54,  // IDAT chunk
                0x08, 0xD7, 0x63, 0x60, 0x60, 0x00, 0x00, 0x00,  // Compressed pixel data
                0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33,        // End of IDAT + CRC
                0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,  // IEND chunk
                0xAE, 0x42, 0x60, 0x82                           // CRC for IEND
            };

            File.WriteAllBytes(path, pngData);
        }

        #endregion
    }
}
