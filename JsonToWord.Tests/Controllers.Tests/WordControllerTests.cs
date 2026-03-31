using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord;
using JsonToWord.Controllers;
using JsonToWord.Models;
using JsonToWord.Models.S3;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Moq;
using Newtonsoft.Json.Linq;

namespace JsonToWord.Controllers.Tests
{
    public class WordControllerTests
    {
        [Fact]
        public void GetStatus_ReturnsOk()
        {
            var controller = new WordController(
                new Mock<IAWSS3Service>().Object,
                new Mock<IWordService>().Object,
                new Mock<ILogger<WordController>>().Object);

            var result = controller.GetStatus();

            var ok = Assert.IsType<OkObjectResult>(result);
            Assert.Contains("Online", ok.Value?.ToString());
        }

        [Fact]
        public async Task CreateWordDocument_DirectDownload_ReturnsDownloadableFile()
        {
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

                var awsService = new Mock<IAWSS3Service>();
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.IsAny<Uri>(), It.IsAny<string>()))
                    .Returns(templatePath);

                var downloadable = new DownloadableObjectModel
                {
                    FileName = "template.docx",
                    Base64 = Convert.ToBase64String(new byte[] { 1, 2, 3 }),
                    ApplicationType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                };

                var wordService = new Mock<IWordService>();
                var callOrder = new List<string>();
                wordService.Setup(s => s.Create(It.IsAny<WordModel>())).Returns(templatePath);
                wordService
                    .Setup(s => s.CreateDownloadableFile(templatePath))
                    .Callback(() => callOrder.Add("download"))
                    .Returns(downloadable);
                awsService
                    .Setup(s => s.CleanUp(templatePath))
                    .Callback(() => callOrder.Add("cleanup"));

                var controller = new WordController(
                    awsService.Object,
                    wordService.Object,
                    new Mock<ILogger<WordController>>().Object);

                var payload = JObject.FromObject(new
                {
                    TemplatePath = "https://example.com/template.docx",
                    UploadProperties = new { FileName = "template.docx", EnableDirectDownload = true },
                    ContentControls = new object[0],
                    FormattingSettings = new { ProcessVoidList = false }
                });

                var result = await controller.CreateWordDocument(payload);

                var ok = Assert.IsType<OkObjectResult>(result);
                Assert.Same(downloadable, ok.Value);
                wordService.Verify(s => s.Create(It.IsAny<WordModel>()), Times.Once);
                awsService.Verify(s => s.CleanUp(templatePath), Times.Once);
                Assert.Equal(new[] { "download", "cleanup" }, callOrder);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task CreateWordDocument_WithoutTemplatePath_BuildsTemporaryTemplateFromContentControls()
        {
            var awsService = new Mock<IAWSS3Service>();
            WordModel capturedModel = null;

            var downloadable = new DownloadableObjectModel
            {
                FileName = "historical-report.docx",
                Base64 = Convert.ToBase64String(new byte[] { 1, 2, 3 }),
                ApplicationType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            };

            var wordService = new Mock<IWordService>();
            wordService
                .Setup(s => s.Create(It.IsAny<WordModel>()))
                .Callback<WordModel>(model =>
                {
                    capturedModel = model;
                    Assert.False(string.IsNullOrWhiteSpace(model?.LocalPath));
                    Assert.True(File.Exists(model.LocalPath));
                })
                .Returns<WordModel>(model => model.LocalPath);
            wordService
                .Setup(s => s.CreateDownloadableFile(It.IsAny<string>()))
                .Returns(downloadable);

            var controller = new WordController(
                awsService.Object,
                wordService.Object,
                new Mock<ILogger<WordController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "historical-report.docx", EnableDirectDownload = true },
                ContentControls = new[]
                {
                    new { Title = "historical-compare-report-content-control", WordObjects = new object[0] }
                },
                FormattingSettings = new { ProcessVoidList = false }
            });

            var result = await controller.CreateWordDocument(payload);

            var ok = Assert.IsType<OkObjectResult>(result);
            Assert.Same(downloadable, ok.Value);
            Assert.NotNull(capturedModel);
            Assert.Equal("historical-report.docx", Path.GetFileName(capturedModel.LocalPath));
            awsService.Verify(s => s.DownloadFileFromS3BucketAsync(It.IsAny<Uri>(), It.IsAny<string>()), Times.Never);
            wordService.Verify(s => s.CreateDownloadableFile(capturedModel.LocalPath), Times.Once);
            awsService.Verify(s => s.CleanUp(capturedModel.LocalPath), Times.Once);
        }

        [Fact]
        public async Task CreateWordDocument_WithoutTemplatePath_UsesRequestedFileNameForUpload()
        {
            var awsService = new Mock<IAWSS3Service>();
            WordModel capturedModel = null;
            UploadProperties capturedUploadProperties = null;

            var wordService = new Mock<IWordService>();
            wordService
                .Setup(s => s.Create(It.IsAny<WordModel>()))
                .Callback<WordModel>(model =>
                {
                    capturedModel = model;
                    Assert.False(string.IsNullOrWhiteSpace(model?.LocalPath));
                    Assert.True(File.Exists(model.LocalPath));
                })
                .Returns<WordModel>(model => model.LocalPath);

            awsService
                .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                .Callback<UploadProperties>(props => capturedUploadProperties = props)
                .ReturnsAsync(new AWSUploadResult<string> { Status = true, Data = "https://minio.example/historical-report.docx" });

            var controller = new WordController(
                awsService.Object,
                wordService.Object,
                new Mock<ILogger<WordController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new
                {
                    FileName = "historical-report.docx",
                    EnableDirectDownload = false,
                    BucketName = "bucket"
                },
                ContentControls = new[]
                {
                    new { Title = "historical-compare-report-content-control", WordObjects = new object[0] }
                },
                FormattingSettings = new { ProcessVoidList = false }
            });

            var result = await controller.CreateWordDocument(payload);

            var ok = Assert.IsType<OkObjectResult>(result);
            Assert.Equal("https://minio.example/historical-report.docx", ok.Value);
            Assert.NotNull(capturedModel);
            Assert.NotNull(capturedUploadProperties);
            Assert.Equal("historical-report.docx", Path.GetFileName(capturedModel.LocalPath));
            Assert.Equal("historical-report.docx", Path.GetFileName(capturedUploadProperties.LocalFilePath));
            awsService.Verify(s => s.CleanUp(capturedModel.LocalPath), Times.Once);
        }

        [Fact]
        public async Task CreateWordDocument_TimeMachineHtmlPayload_RendersHtmlToWordDocument()
        {
            var awsService = new Mock<IAWSS3Service>();
            var controller = new WordController(
                awsService.Object,
                CreateRealWordService(),
                new Mock<ILogger<WordController>>().Object);

            const string reportHtml = @"
                <html>
                    <body>
                        <h1>Time machine summary</h1>
                        <p>Snapshot comparison generated from historical query data.</p>
                        <table>
                            <tr><th>Work Item</th><th>Status</th></tr>
                            <tr><td>Work Item 101</td><td>Added</td></tr>
                            <tr><td>Work Item 202</td><td>Removed</td></tr>
                        </table>
                    </body>
                </html>";

            var payload = JObject.FromObject(new
            {
                UploadProperties = new
                {
                    FileName = "time-machine-report.docx",
                    EnableDirectDownload = true
                },
                ContentControls = new[]
                {
                    new
                    {
                        Title = "historical-compare-report-content-control",
                        WordObjects = new[]
                        {
                            new
                            {
                                type = "Html",
                                html = reportHtml,
                                font = "Calibri",
                                fontSize = 11
                            }
                        }
                    }
                },
                FormattingSettings = new
                {
                    ProcessVoidList = false
                }
            });

            var result = await controller.CreateWordDocument(payload);

            var ok = Assert.IsType<OkObjectResult>(result);
            var downloadable = Assert.IsType<DownloadableObjectModel>(ok.Value);
            Assert.Equal("time-machine-report.docx", downloadable.FileName);
            Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", downloadable.ApplicationType);

            var generatedDocumentBytes = Convert.FromBase64String(downloadable.Base64);
            using var stream = new MemoryStream(generatedDocumentBytes);
            using var generatedDocument = WordprocessingDocument.Open(stream, false);
            var renderedText = string.Join(" ", generatedDocument.MainDocumentPart.Document.Body.Descendants<Text>().Select(t => t.Text));

            Assert.Contains("Work Item 101", renderedText);
            Assert.Contains("Added", renderedText);
            Assert.Contains("Removed", renderedText);
            Assert.Contains("Snapshot comparison generated from historical query data.", renderedText);
            awsService.Verify(s => s.DownloadFileFromS3BucketAsync(It.IsAny<Uri>(), It.IsAny<string>()), Times.Never);
        }

        [Fact]
        public void CreateWordDocumentByFile_ReturnsOk()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var jsonPath = Path.Combine(tempDir, "payload.json");

            try
            {
                var modelJson = JObject.FromObject(new
                {
                    TemplatePath = "https://example.com/template.docx",
                    UploadProperties = new { FileName = "template.docx" },
                    ContentControls = new object[0]
                });
                File.WriteAllText(jsonPath, modelJson.ToString());

                var awsService = new Mock<IAWSS3Service>();
                var wordService = new Mock<IWordService>();
                wordService.Setup(s => s.Create(It.IsAny<WordModel>())).Returns("output.docx");

                var controller = new WordController(
                    awsService.Object,
                    wordService.Object,
                    new Mock<ILogger<WordController>>().Object);

                var payload = JObject.FromObject(new { jsonFilePath = jsonPath });

                var result = controller.CreateWordDocumentByFile(payload);

                var ok = Assert.IsType<OkObjectResult>(result);
                Assert.Equal("output.docx", ok.Value);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task CreateWordDocument_UploadsAndCleansAttachments()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var templatePath = Path.Combine(tempDir, "template.docx");
            var listJsonPath = Path.Combine(tempDir, "cc-list.json");
            var singleJsonPath = Path.Combine(tempDir, "cc-single.json");
            var attachmentPath = Path.Combine(tempDir, "attachment.bin");
            var documentPath = Path.Combine(tempDir, "output.docx");

            try
            {
                using (var doc = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("content")))));
                }

                File.WriteAllText(listJsonPath, "[{\"Title\":\"cc1\",\"WordObjects\":[]}]");
                File.WriteAllText(singleJsonPath, "{\"Title\":\"cc2\",\"WordObjects\":[]}");
                File.WriteAllText(attachmentPath, "attachment");

                var awsService = new Mock<IAWSS3Service>();
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.Is<Uri>(u => u.ToString() == "https://example.com/cc-list.json"), "cc-list.json"))
                    .Returns(listJsonPath);
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.Is<Uri>(u => u.ToString() == "https://example.com/cc-single.json"), "cc-single.json"))
                    .Returns(singleJsonPath);
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.Is<Uri>(u => u.ToString() == "https://example.com/template.docx"), "template.docx"))
                    .Returns(templatePath);
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.Is<Uri>(u => u.ToString() == "https://example.com/attachment.bin"), "attachment.bin"))
                    .Returns(attachmentPath);
                awsService
                    .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                    .ReturnsAsync(new AWSUploadResult<string> { Status = true, Data = "https://minio.example/output.docx" });

                WordModel capturedModel = null;
                var wordService = new Mock<IWordService>();
                wordService
                    .Setup(s => s.Create(It.IsAny<WordModel>()))
                    .Callback<WordModel>(model => capturedModel = model)
                    .Returns(documentPath);

                var controller = new WordController(
                    awsService.Object,
                    wordService.Object,
                    new Mock<ILogger<WordController>>().Object);

                var payload = JObject.FromObject(new
                {
                    TemplatePath = "https://example.com/template.docx",
                    UploadProperties = new { FileName = "template.docx", EnableDirectDownload = false, BucketName = "bucket" },
                    JsonDataList = new[]
                    {
                        new { JsonPath = "https://example.com/cc-list.json", JsonName = "cc-list.json" },
                        new { JsonPath = "https://example.com/cc-single.json", JsonName = "cc-single.json" }
                    },
                    MinioAttachmentData = new[]
                    {
                        new { attachmentMinioPath = "https://example.com/attachment.bin", minioFileName = "attachment.bin" }
                    }
                });

                var result = await controller.CreateWordDocument(payload);

                var ok = Assert.IsType<OkObjectResult>(result);
                Assert.Equal("https://minio.example/output.docx", ok.Value);
                Assert.NotNull(capturedModel);
                Assert.Equal(2, capturedModel.ContentControls.Count);
                Assert.Equal(documentPath, capturedModel.UploadProperties.LocalFilePath);
                awsService.Verify(s => s.CleanUp(listJsonPath), Times.Once);
                awsService.Verify(s => s.CleanUp(singleJsonPath), Times.Once);
                awsService.Verify(s => s.CleanUp(templatePath), Times.Once);
                awsService.Verify(s => s.CleanUp(attachmentPath), Times.Once);
                awsService.Verify(s => s.CleanUp(documentPath), Times.Once);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task CreateWordDocument_UploadFails_ReturnsStatusCode()
        {
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

                var awsService = new Mock<IAWSS3Service>();
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.IsAny<Uri>(), It.IsAny<string>()))
                    .Returns(templatePath);
                awsService
                    .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                    .ReturnsAsync(new AWSUploadResult<string> { Status = false, StatusCode = 500 });

                var wordService = new Mock<IWordService>();
                wordService.Setup(s => s.Create(It.IsAny<WordModel>())).Returns(Path.Combine(tempDir, "output.docx"));

                var controller = new WordController(
                    awsService.Object,
                    wordService.Object,
                    new Mock<ILogger<WordController>>().Object);

                var payload = JObject.FromObject(new
                {
                    TemplatePath = "https://example.com/template.docx",
                    UploadProperties = new { FileName = "template.docx", EnableDirectDownload = false, BucketName = "bucket" }
                });

                var result = await controller.CreateWordDocument(payload);

                var status = Assert.IsType<StatusCodeResult>(result);
                Assert.Equal(500, status.StatusCode);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void CreateWordDocumentByFile_ReturnsBadRequest_ForMissingFile()
        {
            var awsService = new Mock<IAWSS3Service>();
            var wordService = new Mock<IWordService>();
            var controller = new WordController(
                awsService.Object,
                wordService.Object,
                new Mock<ILogger<WordController>>().Object);

            var payload = JObject.FromObject(new { jsonFilePath = "missing.json" });

            var result = controller.CreateWordDocumentByFile(payload);

            Assert.IsType<BadRequestObjectResult>(result);
        }

        private static IWordService CreateRealWordService()
        {
            var contentControlService = new ContentControlService(
                new Mock<ILogger<ContentControlService>>().Object,
                new DocumentValidatorService(new Mock<ILogger<DocumentValidatorService>>().Object));

            var pictureService = new Mock<IPictureService>();
            var htmlService = new HtmlService(
                contentControlService,
                new DocumentValidatorService(new Mock<ILogger<DocumentValidatorService>>().Object),
                new Mock<ILogger<HtmlService>>().Object,
                pictureService.Object,
                new ParagraphService());

            return new WordService(
                contentControlService,
                Mock.Of<ITableService>(),
                pictureService.Object,
                Mock.Of<ITextService>(),
                htmlService,
                Mock.Of<IFileService>(),
                Mock.Of<IVoidListService>(),
                new DocumentService(new Mock<ILogger<DocumentService>>().Object),
                new SectionPlaceholderService(new Mock<ILogger<SectionPlaceholderService>>().Object),
                new Mock<ILogger<WordService>>().Object);
        }
    }
}
