using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ICSharpCode.SharpZipLib.Zip;
using JsonToWord;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    [Collection("NonParallel")]
    public class WordServiceTests
    {
        [Fact]
        public void CreateDownloadableFile_ReturnsBase64AndMimeType()
        {
            var service = CreateService();

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, "file.docx");
            var bytes = new byte[] { 1, 2, 3, 4 };
            File.WriteAllBytes(filePath, bytes);

            try
            {
                var result = service.CreateDownloadableFile(filePath);

                Assert.Equal("file.docx", result.FileName);
                Assert.Equal(Convert.ToBase64String(bytes), result.Base64);
                Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", result.ApplicationType);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Theory]
        [InlineData("file.pdf", "application/pdf")]
        [InlineData("file.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
        [InlineData("file.bin", "application/octet-stream")]
        public void CreateDownloadableFile_ReturnsMimeTypeForExtensions(string fileName, string expectedType)
        {
            var service = CreateService();

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, fileName);
            File.WriteAllBytes(filePath, new byte[] { 9, 8, 7 });

            try
            {
                var result = service.CreateDownloadableFile(filePath);

                Assert.Equal(expectedType, result.ApplicationType);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void Create_WithNoContentControls_ReturnsDocumentPath()
        {
            var mocks = CreateServiceWithMocks();
            var service = mocks.Service;

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var templatePath = Path.Combine(tempDir, "template.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));
                }

                mocks.DocumentService
                    .Setup(d => d.CreateDocument(templatePath))
                    .Returns(templatePath);

                var model = new WordModel
                {
                    LocalPath = templatePath,
                    ContentControls = new List<WordContentControl>(),
                    FormattingSettings = new FormattingSettings { ProcessVoidList = false }
                };

                var resultPath = service.Create(model);

                Assert.Equal(templatePath, resultPath);
                mocks.DocumentService.Verify(d => d.CreateDocument(templatePath), Times.Once);
                mocks.ContentControlService.Verify(c => c.ClearContentControlHeadingMap(), Times.Once);
                mocks.VoidListService.Verify(v => v.CreateVoidList(It.IsAny<string>()), Times.Never);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void Create_WithNullFormattingSettings_DoesNotThrow_AndSkipsVoidList()
        {
            var mocks = CreateServiceWithMocks();
            var service = mocks.Service;

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var templatePath = Path.Combine(tempDir, "template.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));
                }

                mocks.DocumentService
                    .Setup(d => d.CreateDocument(templatePath))
                    .Returns(templatePath);

                var model = new WordModel
                {
                    LocalPath = templatePath,
                    ContentControls = new List<WordContentControl>(),
                    FormattingSettings = null
                };

                var resultPath = service.Create(model);

                Assert.Equal(templatePath, resultPath);
                mocks.VoidListService.Verify(v => v.CreateVoidList(It.IsAny<string>()), Times.Never);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void Create_WithVoidList_CreatesZipWithVoidListFile()
        {
            var mocks = CreateServiceWithMocks();
            var service = mocks.Service;

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;

            var templatePath = Path.Combine(tempDir, "template-2025-08-21-16:39:15.docx");
            var voidListPath = Path.Combine(tempDir, "void-list.txt");

            try
            {
                using (var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));
                }

                File.WriteAllText(voidListPath, "void-list");

                mocks.DocumentService
                    .Setup(d => d.CreateDocument(templatePath))
                    .Returns(templatePath);
                mocks.VoidListService
                    .Setup(v => v.CreateVoidList(templatePath))
                    .Returns(new List<string> { voidListPath });

                var model = new WordModel
                {
                    LocalPath = templatePath,
                    ContentControls = new List<WordContentControl>(),
                    FormattingSettings = new FormattingSettings { ProcessVoidList = true }
                };

                var resultPath = service.Create(model);

                Assert.EndsWith(".zip", resultPath);
                Assert.True(File.Exists(resultPath));

                using var zip = new ZipFile(resultPath);
                Assert.NotNull(zip.GetEntry(Path.GetFileName(templatePath).Replace(":", "_")));
                Assert.NotNull(zip.GetEntry(Path.GetFileName(voidListPath)));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void Create_WithNonOfficeAttachment_ZipsWithAttachments()
        {
            var mocks = CreateServiceWithMocks();
            var service = mocks.Service;

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;

            var templatePath = Path.Combine(tempDir, "template.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));
                }

                var sdtBlock = new SdtBlock(new SdtContentBlock());
                mocks.ContentControlService
                    .Setup(c => c.FindContentControl(It.IsAny<WordprocessingDocument>(), It.IsAny<string>()))
                    .Returns(sdtBlock);
                mocks.ContentControlService
                    .Setup(c => c.IsUnderStandardHeading(It.IsAny<SdtBlock>()))
                    .Returns(true);

                mocks.DocumentService
                    .Setup(d => d.CreateDocument(templatePath))
                    .Returns(templatePath);

                mocks.FileService
                    .Setup(f => f.Insert(It.IsAny<WordprocessingDocument>(), It.IsAny<string>(), It.IsAny<WordAttachment>()))
                    .Callback(() =>
                    {
                        Directory.CreateDirectory("attachments");
                        var attachmentPath = Path.Combine("attachments", "file.txt");
                        File.WriteAllText(attachmentPath, "attachment");
                        mocks.FileService.Raise(m => m.nonOfficeAttachmentEventHandler += null);
                    });

                var model = new WordModel
                {
                    LocalPath = templatePath,
                    ContentControls = new List<WordContentControl>
                    {
                        new WordContentControl
                        {
                            Title = "cc1",
                            WordObjects = new List<IWordObject>
                            {
                                new WordAttachment { Type = WordObjectType.File, Name = "file.txt", Path = "file.txt" }
                            }
                        }
                    },
                    FormattingSettings = new FormattingSettings()
                };

                var resultPath = service.Create(model);

                Assert.EndsWith(".zip", resultPath);
                Assert.True(File.Exists(resultPath));

                using var zip = new ZipFile(resultPath);
                Assert.True(zip.Cast<ZipEntry>().Any(entry => entry.Name.EndsWith("file.txt", StringComparison.Ordinal)));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void Create_WithContentControls_CallsHandlersForEachObjectType()
        {
            var mocks = CreateServiceWithMocks();
            var service = mocks.Service;

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var templatePath = Path.Combine(tempDir, "template.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));
                }

                mocks.DocumentService
                    .Setup(d => d.CreateDocument(templatePath))
                    .Returns(templatePath);

                mocks.ContentControlService
                    .Setup(c => c.FindContentControl(It.IsAny<WordprocessingDocument>(), It.IsAny<string>()))
                    .Returns(new SdtBlock(new SdtContentBlock()));
                mocks.ContentControlService
                    .Setup(c => c.IsUnderStandardHeading(It.IsAny<SdtBlock>()))
                    .Returns(false);
                mocks.ContentControlService
                    .Setup(c => c.GetContentControlHeadingStatus(It.IsAny<string>()))
                    .Returns(true);

                var contentControl = new WordContentControl
                {
                    Title = "cc1",
                    ForceClean = true,
                    WordObjects = new List<IWordObject>
                    {
                        new WordAttachment { Type = WordObjectType.File, Name = "file", Path = "file.docx" },
                        new WordHtml { Type = WordObjectType.Html, Html = "<p>Hi</p>" },
                        new WordAttachment { Type = WordObjectType.Picture, Name = "pic", Path = "pic.png" },
                        new WordParagraph { Type = WordObjectType.Paragraph, Runs = new List<WordRun>() },
                        new WordTable { Type = WordObjectType.Table, Rows = new List<WordTableRow>() }
                    }
                };

                var model = new WordModel
                {
                    LocalPath = templatePath,
                    ContentControls = new List<WordContentControl>
                    {
                        new WordContentControl { Title = null, WordObjects = new List<IWordObject>() },
                        contentControl
                    },
                    FormattingSettings = new FormattingSettings()
                };

                var resultPath = service.Create(model);

                Assert.Equal(templatePath, resultPath);
                mocks.FileService.Verify(f => f.Insert(It.IsAny<WordprocessingDocument>(), "cc1", It.IsAny<WordAttachment>()), Times.Once);
                mocks.PictureService.Verify(p => p.Insert(It.IsAny<WordprocessingDocument>(), "cc1", It.IsAny<WordAttachment>()), Times.Once);
                mocks.HtmlService.Verify(h => h.Insert(It.IsAny<WordprocessingDocument>(), "cc1", It.IsAny<WordHtml>(), It.IsAny<FormattingSettings>()), Times.Once);
                mocks.TextService.Verify(t => t.Write(It.IsAny<WordprocessingDocument>(), "cc1", It.IsAny<WordParagraph>(), It.IsAny<bool>()), Times.Once);
                mocks.TableService.Verify(t => t.Insert(It.IsAny<WordprocessingDocument>(), "cc1", It.IsAny<WordTable>(), It.IsAny<FormattingSettings>()), Times.Once);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        private static WordService CreateService()
        {
            return CreateServiceWithMocks().Service;
        }

        private static (WordService Service,
            Mock<IContentControlService> ContentControlService,
            Mock<IFileService> FileService,
            Mock<IVoidListService> VoidListService,
            Mock<IDocumentService> DocumentService,
            Mock<ITableService> TableService,
            Mock<IPictureService> PictureService,
            Mock<ITextService> TextService,
            Mock<IHtmlService> HtmlService) CreateServiceWithMocks()
        {
            var contentControlService = new Mock<IContentControlService>();
            var tableService = new Mock<ITableService>();
            var pictureService = new Mock<IPictureService>();
            var textService = new Mock<ITextService>();
            var htmlService = new Mock<IHtmlService>();
            var fileService = new Mock<IFileService>();
            var voidListService = new Mock<IVoidListService>();
            var documentService = new Mock<IDocumentService>();
            var logger = new Mock<ILogger<WordService>>();

            var service = new WordService(
                contentControlService.Object,
                tableService.Object,
                pictureService.Object,
                textService.Object,
                htmlService.Object,
                fileService.Object,
                voidListService.Object,
                documentService.Object,
                logger.Object);

            return (service, contentControlService, fileService, voidListService, documentService, tableService, pictureService, textService, htmlService);
        }
    }
}
