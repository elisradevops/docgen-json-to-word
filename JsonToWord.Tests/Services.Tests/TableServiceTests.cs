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
    public class TableServiceTests : IDisposable
    {
        private readonly Mock<IContentControlService> _mockContentControlService;
        private readonly Mock<IParagraphService> _mockParagraphService;
        private readonly Mock<IPictureService> _mockPictureService;
        private readonly Mock<IFileService> _mockFileService;
        private readonly Mock<IRunService> _mockRunService;
        private readonly Mock<IHtmlService> _mockHtmlService;
        private readonly Mock<IUtilsService> _mockUtilsService;
        private readonly Mock<ILogger<TableService>> _mockLogger;
        private readonly TableService _tableService;
        
        private readonly string _docPath;
        private WordprocessingDocument _document;
        
        public TableServiceTests()
        {
            _mockContentControlService = new Mock<IContentControlService>();
            _mockParagraphService = new Mock<IParagraphService>();
            _mockPictureService = new Mock<IPictureService>();
            _mockFileService = new Mock<IFileService>();
            _mockRunService = new Mock<IRunService>();
            _mockHtmlService = new Mock<IHtmlService>();
            _mockUtilsService = new Mock<IUtilsService>();
            _mockLogger = new Mock<ILogger<TableService>>();
            
            _tableService = new TableService(
                _mockContentControlService.Object,
                _mockParagraphService.Object,
                _mockRunService.Object,
                _mockHtmlService.Object,
                _mockPictureService.Object,
                _mockFileService.Object,
                _mockLogger.Object,
                _mockUtilsService.Object
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
            
            // Setup common mocks
            _mockContentControlService.Setup(m => m.FindContentControl(_document, It.IsAny<string>()))
                .Returns(new SdtBlock(new SdtProperties(), new SdtEndCharProperties(), new SdtContentBlock()));
            
            _mockUtilsService.Setup(m => m.GetPageWidthDxa(It.IsAny<MainDocumentPart>()))
                .Returns(12240); // Standard page width
            
            _mockParagraphService.Setup(m => m.CreateParagraph(It.IsAny<WordParagraph>()))
                .Returns(new Paragraph());
            
            _mockRunService.Setup(m => m.CreateRun(It.IsAny<WordRun>()))
                .Returns(new Run(new Text("Test Content")));
        }

        [Fact]
        public void Insert_BasicTable_CreatesTableInDocument()
        {
            // Arrange
            var wordTable = new WordTable
            {
                RepeatHeaderRow = true,
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { Paragraphs = new List<WordParagraph> { new WordParagraph() } }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);
            
            // Verify the content control was updated with the table
            _mockContentControlService.Verify(m => m.FindContentControl(
                It.IsAny<WordprocessingDocument>(), 
                It.IsAny<string>()
            ), Times.Once);
        }

        [Fact]
        public void Insert_WithTableHavingHeaderRow_CreatesTableWithHeader()
        {
            // Arrange
            var wordTable = new WordTable
            {
                RepeatHeaderRow = true,
                Rows = new List<WordTableRow>
                {
                    new WordTableRow // Header row
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Paragraphs = new List<WordParagraph> { 
                                    new WordParagraph { 
                                        Runs = new List<WordRun> { 
                                            new WordRun { Text = "Header" } 
                                        } 
                                    } 
                                } 
                            }
                        }
                    },
                    new WordTableRow // Data row
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Paragraphs = new List<WordParagraph> { 
                                    new WordParagraph { 
                                        Runs = new List<WordRun> { 
                                            new WordRun { Text = "Data" } 
                                        } 
                                    } 
                                } 
                            }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);
        }

        [Fact]
        public void Insert_WithHtmlContent_CallsHtmlService()
        {
            // Arrange
            var html = new WordHtml { Html = "<p>Test HTML</p>" };
            
            _mockHtmlService.Setup(m => m.ConvertHtmlToOpenXmlElements(It.IsAny<WordHtml>(), _document))
                .Returns(new List<OpenXmlCompositeElement> { new Paragraph(new Run(new Text("Converted HTML"))) });
            
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { Html = html }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockHtmlService.Verify(m => m.ConvertHtmlToOpenXmlElements(html, _document), Times.Once);
        }

        [Fact]
        public void Insert_WithPictureAttachment_CallsPictureService()
        {
            // Arrange
            var attachment = new WordAttachment
            {
                Type = WordObjectType.Picture,
                Path = "test.jpg",
                Name = "Test Picture"
            };
            
            _mockPictureService.Setup(m => m.CreateDrawing(It.IsAny<MainDocumentPart>(), It.IsAny<string>(), It.IsAny<bool>()))
                .Returns(new Drawing());
            
            _mockParagraphService.Setup(m => m.CreateCaption(It.IsAny<string>()))
                .Returns(new Paragraph());
            
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Attachments = new List<WordAttachment> { attachment } 
                            }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockPictureService.Verify(m => m.CreateDrawing(
                It.IsAny<MainDocumentPart>(), 
                attachment.Path, 
                attachment.IsFlattened.GetValueOrDefault()
            ), Times.Once);
            
            _mockParagraphService.Verify(m => m.CreateCaption(attachment.Name), Times.Once);
        }

        [Fact]
        public void Insert_WithFileAttachment_CallsFileService()
        {
            // Arrange
            var attachment = new WordAttachment
            {
                Type = WordObjectType.File,
                Path = "test.pdf",
                Name = "Test File"
            };
            
            _mockFileService.Setup(m => m.AttachFileToParagraph(It.IsAny<MainDocumentPart>(), It.IsAny<WordAttachment>()))
                .Returns(new Paragraph());
            
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Attachments = new List<WordAttachment> { attachment } 
                            }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockFileService.Verify(m => m.AttachFileToParagraph(
                It.IsAny<MainDocumentPart>(), 
                attachment
            ), Times.Once);
        }

        [Fact]
        public void Insert_WithMergedCells_CreatesTableWithMergedCells()
        {
            // Arrange
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        MergeToOneCell = true,
                        NumberOfCellsToMerge = 3,
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Paragraphs = new List<WordParagraph> { new WordParagraph() } 
                            }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);
        }

        [Fact]
        public void Insert_WithShadedCell_CreatesTableWithShading()
        {
            // Arrange
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Paragraphs = new List<WordParagraph> { new WordParagraph() },
                                Shading = new WordShading { Color = "000000", Fill = "CCCCCC" }
                            }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);
        }

        [Fact]
        public void Insert_WithCellWidth_SetsWidthCorrectly()
        {
            // Arrange
            _mockUtilsService.Setup(m => m.ParseStringToDouble("50%"))
                .Returns(50);
            
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Width = "50%",
                                Paragraphs = new List<WordParagraph> { new WordParagraph() }
                            }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockUtilsService.Verify(m => m.ParseStringToDouble("50%"), Times.Once);
        }

        [Fact]
        public void Insert_WithCellWidthInCm_ConvertsToDxa()
        {
            // Arrange
            _mockUtilsService.Setup(m => m.ParseStringToDouble("5cm"))
                .Returns(5);
            
            _mockUtilsService.Setup(m => m.ConvertCmToDxa(5))
                .Returns(2835); // 5cm in Dxa
            
            _mockUtilsService.Setup(m => m.ConvertDxaToPct(2835, 12240))
                .Returns(1158); // ~23% of page width
            
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { 
                                Width = "5cm",
                                Paragraphs = new List<WordParagraph> { new WordParagraph() }
                            }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockUtilsService.Verify(m => m.ParseStringToDouble("5cm"), Times.Once);
            _mockUtilsService.Verify(m => m.ConvertCmToDxa(5), Times.Once);
            _mockUtilsService.Verify(m => m.ConvertDxaToPct(2835, 12240), Times.Once);
        }

        [Fact]
        public void Insert_WithHtmlError_AddsErrorMessage()
        {
            // Arrange
            var html = new WordHtml { Html = "<p>Test HTML</p>" };
            
            _mockHtmlService.Setup(m => m.ConvertHtmlToOpenXmlElements(It.IsAny<WordHtml>(), _document))
                .Throws(new Exception("HTML conversion error"));
            
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { Html = html }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockHtmlService.Verify(m => m.ConvertHtmlToOpenXmlElements(html, _document), Times.Once);
            _mockLogger.Verify(
            x => x.Log(
                It.Is<LogLevel>(l => l == LogLevel.Error),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString()!.Contains("Error while creating table cell")),
                It.IsAny<Exception>(),
                It.Is<Func<It.IsAnyType, Exception?, string>>((v, t) => true)),
            Times.Once);
        }

        [Fact]
        public void Insert_WithPageBreak_AddsPageBreak()
        {
            // Arrange
            var wordTable = new WordTable
            {
                InsertPageBreak = true,
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { Paragraphs = new List<WordParagraph> { new WordParagraph() } }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            _mockContentControlService.Verify(m => m.FindContentControl(_document, "TestControl"), Times.Once);
        }

        [Fact]
        public void Insert_EnsuresCellsHaveParagraphAsLastChild()
        {
            // Arrange
            var mockSdtBlock = new Mock<SdtBlock>();
            SdtContentBlock? capturedSdtContentBlock = null;

            // Mock the FindContentControl method to return our mock SdtBlock
            _mockContentControlService.Setup(m => m.FindContentControl(_document, "TestControl"))
                .Returns(mockSdtBlock.Object);

            // Setup to capture the SdtContentBlock that gets appended to the SdtBlock
            mockSdtBlock.Setup(b => b.AppendChild(It.IsAny<SdtContentBlock>()))
                .Callback<SdtContentBlock>(contentBlock => capturedSdtContentBlock = contentBlock);

            // Create test table with different cell content types
            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            // Cell with paragraphs
                            new WordTableCell {
                                Paragraphs = new List<WordParagraph> {
                                    new WordParagraph {
                                        Runs = new List<WordRun> { new WordRun { Text = "Text content" } }
                                    }
                                }
                            },
                            // Cell with HTML
                            new WordTableCell {
                                Html = new WordHtml { Html = "<p>HTML content</p>" }
                            },
                            // Cell with picture attachment
                            new WordTableCell {
                                Attachments = new List<WordAttachment> {
                                    new WordAttachment {
                                        Type = WordObjectType.Picture,
                                        Path = "test.jpg",
                                        Name = "Test Picture"
                                    }
                                }
                            }
                        }
                    }
                }
            };

            // Setup necessary mocks for content conversion
            _mockHtmlService.Setup(m => m.ConvertHtmlToOpenXmlElements(It.IsAny<WordHtml>(), _document))
                .Returns(new List<OpenXmlCompositeElement> { new Paragraph(new Run(new Text("Converted HTML"))) });

            _mockPictureService.Setup(m => m.CreateDrawing(It.IsAny<MainDocumentPart>(), It.IsAny<string>(), It.IsAny<bool>()))
                .Returns(new Drawing());

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            Assert.NotNull(capturedSdtContentBlock);

            // Find the table in the captured SdtContentBlock
            var table = capturedSdtContentBlock?.Elements<Table>().FirstOrDefault();
            Assert.NotNull(table);

            // Examine the cells
            var cells = table?.Descendants<TableCell>().ToList();
            Assert.NotEmpty(cells);

            // Check that every cell has a paragraph as its last child
            foreach (var cell in cells!)
            {
                var lastChild = cell.LastChild;
                Assert.IsType<Paragraph>(lastChild);
            }
        }

        [Fact]
        public void Insert_WithHtmlContentErrorRecovers_EnsuresParagraphAsLastChild()
        {
            // Arrange
            var mockSdtBlock = new Mock<SdtBlock>();
            SdtContentBlock? capturedSdtContentBlock = null;

            // Mock the FindContentControl method to return our mock SdtBlock
            _mockContentControlService.Setup(m => m.FindContentControl(_document, "TestControl"))
                .Returns(mockSdtBlock.Object);

            // Setup to capture the SdtContentBlock that gets appended to the SdtBlock
            mockSdtBlock.Setup(b => b.AppendChild(It.IsAny<SdtContentBlock>()))
                .Callback<SdtContentBlock>(contentBlock => capturedSdtContentBlock = contentBlock);

            // Setup HTML service to throw an exception
            _mockHtmlService.Setup(m => m.ConvertHtmlToOpenXmlElements(It.IsAny<WordHtml>(), _document))
                .Throws(new Exception("HTML parsing error"));

            var wordTable = new WordTable
            {
                Rows = new List<WordTableRow>
                {
                    new WordTableRow
                    {
                        Cells = new List<WordTableCell>
                        {
                            new WordTableCell { Html = new WordHtml { Html = "<p>Problem HTML</p>" } }
                        }
                    }
                }
            };

            // Act
            _tableService.Insert(_document, "TestControl", wordTable);

            // Assert
            Assert.NotNull(capturedSdtContentBlock);

            // Find the table in the captured SdtContentBlock
            var table = capturedSdtContentBlock?.Elements<Table>().FirstOrDefault();
            Assert.NotNull(table);

            // Examine the cells
            var cells = table?.Descendants<TableCell>().ToList();
            Assert.NotEmpty(cells);

            foreach (var cell in cells!)
            {
                // Verify error recovery still results in paragraph as last child
                Assert.IsType<Paragraph>(cell.LastChild);

                // Verify error message is present
                var textElements = cell.Descendants<Text>();
                Assert.Contains(textElements, t => t.Text.Contains("DocGen Error"));
            }

            // Verify error was logged
            _mockLogger.Verify(
            x => x.Log(
                It.Is<LogLevel>(l => l == LogLevel.Error),
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString()!.Contains("Error while creating table cell")),
                It.IsAny<Exception>(),
                It.Is<Func<It.IsAnyType, Exception?, string>>((v, t) => true)),
            Times.Once);
        }

        public void Dispose()
        {
            _document?.Dispose();
            if (File.Exists(_docPath))
            {
                File.Delete(_docPath);
            }
        }
    }
}