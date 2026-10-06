using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
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
    public class ExcelControllerTests
    {
        [Fact]
        public void GetStatus_ReturnsOk()
        {
            var controller = new ExcelController(
                new Mock<IAWSS3Service>().Object,
                new Mock<IExcelService>().Object,
                new Mock<ILogger<ExcelController>>().Object);

            var result = controller.GetStatus();

            var ok = Assert.IsType<OkObjectResult>(result);
            Assert.Contains("Online", ok.Value?.ToString());
        }

        [Fact]
        public async Task CreateExcelDocument_AppendsExtensionAndUploads()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();
            var logger = new Mock<ILogger<ExcelController>>();

            ExcelModel capturedModel = null;
            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Callback<ExcelModel>(model => capturedModel = model)
                .Returns((ExcelModel model) => model.LocalPath);

            awsService
                .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                .ReturnsAsync(new AWSUploadResult<string> { Status = true, Data = "https://minio.example/report.xlsx" });

            var controller = new ExcelController(awsService.Object, excelService.Object, logger.Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report", BucketName = "bucket", Region = "us" }
            });

            var result = await controller.CreateExcelDocument(payload);

            var ok = Assert.IsType<OkObjectResult>(result);
            Assert.Equal("https://minio.example/report.xlsx", ok.Value);
            Assert.NotNull(capturedModel);
            Assert.Equal("report.xlsx", capturedModel.UploadProperties.FileName);
            Assert.EndsWith("report.xlsx", capturedModel.LocalPath);
            Assert.StartsWith("TempFiles" + System.IO.Path.DirectorySeparatorChar + "json-to-word-", capturedModel.LocalPath);
            awsService.Verify(s => s.UploadFileToMinioBucketAsync(It.Is<UploadProperties>(p => p.LocalFilePath == capturedModel.LocalPath)), Times.Once);
            awsService.Verify(s => s.CleanUp(capturedModel.LocalPath), Times.Once);
        }

        [Fact]
        public async Task CreateExcelDocument_SameFileNameTwice_UsesDifferentOutputPaths()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();
            var logger = new Mock<ILogger<ExcelController>>();

            var paths = new System.Collections.Generic.List<string>();
            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Callback<ExcelModel>(model => paths.Add(model.LocalPath))
                .Returns((ExcelModel model) => model.LocalPath);
            awsService
                .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                .ReturnsAsync(new AWSUploadResult<string> { Status = true, Data = "https://minio.example/report.xlsx" });

            var controller = new ExcelController(awsService.Object, excelService.Object, logger.Object);
            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report", BucketName = "bucket", Region = "us" }
            });

            await controller.CreateExcelDocument(payload);
            await controller.CreateExcelDocument(payload);

            Assert.Equal(2, paths.Count);
            Assert.NotEqual(paths[0], paths[1]);
            Assert.All(paths, p => Assert.EndsWith("report.xlsx", p));
        }

        [Fact]
        public async Task CreateExcelDocument_KeepsARequestSuppliedFileNameInsideItsTempDirectory()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();
            ExcelModel capturedModel = null;
            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Callback<ExcelModel>(model => capturedModel = model)
                .Returns((ExcelModel model) => model.LocalPath);
            awsService
                .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                .ReturnsAsync(new AWSUploadResult<string> { Status = true, Data = "https://minio.example/x.xlsx" });
            var controller = new ExcelController(awsService.Object, excelService.Object, new Mock<ILogger<ExcelController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "/etc/cron.d/evil", BucketName = "bucket", Region = "us" }
            });
            await controller.CreateExcelDocument(payload);

            Assert.NotNull(capturedModel);
            Assert.StartsWith("TempFiles" + Path.DirectorySeparatorChar + "json-to-word-", capturedModel.LocalPath);
            Assert.Equal("evil.xlsx", Path.GetFileName(capturedModel.LocalPath));
        }

        [Theory]
        [InlineData(typeof(HttpRequestException), 502)]
        [InlineData(typeof(TaskCanceledException), 504)]
        [InlineData(typeof(InvalidOperationException), 500)]
        public async Task CreateExcelDocument_MapsAFailedOrTimedOutDownloadToAGatewayStatus(Type failure, int expectedStatus)
        {
            var awsService = new Mock<IAWSS3Service>();
            awsService
                .Setup(s => s.DownloadFileFromS3BucketAsync(It.IsAny<Uri>(), It.IsAny<string>()))
                .ThrowsAsync((Exception)Activator.CreateInstance(failure));
            var controller = new ExcelController(awsService.Object, new Mock<IExcelService>().Object, new Mock<ILogger<ExcelController>>().Object);

            var result = await controller.CreateExcelDocument(JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report", BucketName = "b", Region = "us" },
                JsonDataList = new[] { new { JsonPath = "https://example.com/cc.json", JsonName = "cc.json" } }
            }));

            Assert.Equal(expectedStatus, Assert.IsType<ObjectResult>(result).StatusCode);
        }

        [Fact]
        public async Task CreateExcelDocument_RenderFailure_CleansUpEvenWhenNoOutputFileWasWritten()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();
            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Throws(new InvalidOperationException("render failed"));
            var controller = new ExcelController(awsService.Object, excelService.Object, new Mock<ILogger<ExcelController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report", BucketName = "bucket", Region = "us" }
            });
            var result = await controller.CreateExcelDocument(payload);

            Assert.IsType<ObjectResult>(result);
            awsService.Verify(s => s.CleanUp(It.Is<string>(p => p.EndsWith("report.xlsx"))), Times.Once);
        }

        [Fact]
        public async Task CreateExcelDocument_MalformedContentControlJson_StillCleansUpTheDownload()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var brokenPath = Path.Combine(tempDir, "broken.json");
            try
            {
                File.WriteAllText(brokenPath, "[{ not json");
                var awsService = new Mock<IAWSS3Service>();
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.IsAny<Uri>(), "broken.json"))
                    .ReturnsAsync(brokenPath);
                var controller = new ExcelController(awsService.Object, new Mock<IExcelService>().Object, new Mock<ILogger<ExcelController>>().Object);

                var payload = JObject.FromObject(new
                {
                    UploadProperties = new { FileName = "report", BucketName = "bucket", Region = "us" },
                    JsonDataList = new[] { new { JsonPath = "https://example.com/broken.json", JsonName = "broken.json" } }
                });
                var result = await controller.CreateExcelDocument(payload);

                Assert.IsType<ObjectResult>(result);
                awsService.Verify(s => s.CleanUp(brokenPath), Times.Once);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task CreateExcelDocument_ParsesJsonDataList_AndCleansUp()
        {
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var listJsonPath = Path.Combine(tempDir, "cc-list.json");
            var singleJsonPath = Path.Combine(tempDir, "cc-single.json");

            try
            {
                File.WriteAllText(listJsonPath, "[{\"Title\":\"cc1\",\"WordObjects\":[]}]");
                File.WriteAllText(singleJsonPath, "{\"Title\":\"cc2\",\"WordObjects\":[]}");

                var awsService = new Mock<IAWSS3Service>();
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.Is<Uri>(u => u.ToString() == "https://example.com/cc-list.json"), "cc-list.json"))
                    .ReturnsAsync(listJsonPath);
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.Is<Uri>(u => u.ToString() == "https://example.com/cc-single.json"), "cc-single.json"))
                    .ReturnsAsync(singleJsonPath);
                awsService
                    .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                    .ReturnsAsync(new AWSUploadResult<string> { Status = true, Data = "https://minio.example/report.xlsx" });

                ExcelModel capturedModel = null;
                var excelService = new Mock<IExcelService>();
                excelService
                    .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                    .Callback<ExcelModel>(model => capturedModel = model)
                    .Returns((ExcelModel model) => model.LocalPath);

                var controller = new ExcelController(awsService.Object, excelService.Object, new Mock<ILogger<ExcelController>>().Object);

                var payload = JObject.FromObject(new
                {
                    UploadProperties = new { FileName = "report.xlsx", BucketName = "bucket", Region = "us" },
                    JsonDataList = new[]
                    {
                        new { JsonPath = "https://example.com/cc-list.json", JsonName = "cc-list.json" },
                        new { JsonPath = "https://example.com/cc-single.json", JsonName = "cc-single.json" }
                    }
                });

                var result = await controller.CreateExcelDocument(payload);

                var ok = Assert.IsType<OkObjectResult>(result);
                Assert.Equal("https://minio.example/report.xlsx", ok.Value);
                Assert.NotNull(capturedModel);
                Assert.Equal(2, capturedModel.ContentControls.Count);
                Assert.Equal("report.xlsx", capturedModel.UploadProperties.FileName);
                awsService.Verify(s => s.CleanUp(listJsonPath), Times.Once);
                awsService.Verify(s => s.CleanUp(singleJsonPath), Times.Once);
                awsService.Verify(s => s.CleanUp(capturedModel.LocalPath), Times.Once);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task CreateExcelDocument_UploadFails_ReturnsStatusCode()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();

            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Returns("TempFiles/report.xlsx");

            awsService
                .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                .ReturnsAsync(new AWSUploadResult<string> { Status = false, StatusCode = 502 });

            var controller = new ExcelController(awsService.Object, excelService.Object, new Mock<ILogger<ExcelController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report.xlsx", BucketName = "bucket", Region = "us" }
            });

            var result = await controller.CreateExcelDocument(payload);

            var status = Assert.IsType<StatusCodeResult>(result);
            Assert.Equal(502, status.StatusCode);
        }

        [Fact]
        public async Task CreateExcelDocument_JsonDeserializationError_Returns400()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();
            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Throws(new Newtonsoft.Json.JsonReaderException("bad json"));

            var controller = new ExcelController(awsService.Object, excelService.Object, new Mock<ILogger<ExcelController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report.xlsx", BucketName = "bucket" }
            });

            var result = await controller.CreateExcelDocument(payload);

            var objectResult = Assert.IsType<ObjectResult>(result);
            Assert.Equal(400, objectResult.StatusCode);
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(objectResult.Value);
            Assert.Contains("render-document", json);
            Assert.Contains("json-to-word", json);
        }

        [Fact]
        public async Task CreateExcelDocument_InternalRenderingError_Returns500()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();
            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Throws(new InvalidOperationException("render failed"));

            var controller = new ExcelController(awsService.Object, excelService.Object, new Mock<ILogger<ExcelController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report.xlsx", BucketName = "bucket" }
            });

            var result = await controller.CreateExcelDocument(payload);

            var objectResult = Assert.IsType<ObjectResult>(result);
            Assert.Equal(500, objectResult.StatusCode);
        }

        [Fact]
        public async Task CreateExcelZipPackage_InternalError_Returns500()
        {
            var awsService = new Mock<IAWSS3Service>();
            awsService
                .Setup(s => s.UploadFileToMinioBucketAsync(It.IsAny<UploadProperties>()))
                .ThrowsAsync(new InvalidOperationException("zip upload failed"));

            var controller = new ExcelController(awsService.Object, new Mock<IExcelService>().Object, new Mock<ILogger<ExcelController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report.zip", BucketName = "bucket" },
                Files = new[] { new { FileName = "a.xlsx", Base64 = Convert.ToBase64String(new byte[] { 1, 2, 3 }) } }
            });

            var result = await controller.CreateExcelZipPackage(payload);

            var objectResult = Assert.IsType<ObjectResult>(result);
            Assert.Equal(500, objectResult.StatusCode);
        }

        [Fact]
        public async Task CreateExcelDocument_S3Exception_Returns502()
        {
            var awsService = new Mock<IAWSS3Service>();
            var excelService = new Mock<IExcelService>();
            excelService
                .Setup(s => s.CreateExcelDocument(It.IsAny<ExcelModel>()))
                .Throws(new Amazon.S3.AmazonS3Exception("S3 unavailable"));

            var controller = new ExcelController(awsService.Object, excelService.Object, new Mock<ILogger<ExcelController>>().Object);

            var payload = JObject.FromObject(new
            {
                UploadProperties = new { FileName = "report.xlsx", BucketName = "bucket" }
            });

            var result = await controller.CreateExcelDocument(payload);

            var objectResult = Assert.IsType<ObjectResult>(result);
            Assert.Equal(502, objectResult.StatusCode);
        }
    }
}
