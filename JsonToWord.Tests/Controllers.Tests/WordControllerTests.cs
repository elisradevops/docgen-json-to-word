using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Controllers;
using JsonToWord.Models;
using JsonToWord.Models.S3;
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
                wordService.Setup(s => s.Create(It.IsAny<WordModel>())).Returns(templatePath);
                wordService.Setup(s => s.CreateDownloadableFile(templatePath)).Returns(downloadable);

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
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
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
    }
}
