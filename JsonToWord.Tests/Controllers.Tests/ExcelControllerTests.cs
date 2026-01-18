using System;
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
            Assert.EndsWith("TempFiles/report.xlsx", capturedModel.LocalPath);
            awsService.Verify(s => s.UploadFileToMinioBucketAsync(It.Is<UploadProperties>(p => p.LocalFilePath == capturedModel.LocalPath)), Times.Once);
            awsService.Verify(s => s.CleanUp(capturedModel.LocalPath), Times.Once);
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
                    .Returns(listJsonPath);
                awsService
                    .Setup(s => s.DownloadFileFromS3BucketAsync(It.Is<Uri>(u => u.ToString() == "https://example.com/cc-single.json"), "cc-single.json"))
                    .Returns(singleJsonPath);
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
    }
}
