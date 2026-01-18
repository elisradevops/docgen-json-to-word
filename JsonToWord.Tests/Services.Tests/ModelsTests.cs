using System;
using System.Collections.Generic;
using JsonToWord.Models;
using JsonToWord.Models.S3;

namespace JsonToWord.Services.Tests
{
    public class ModelsTests
    {
        [Fact]
        public void ModelProperties_CanRoundTripValues()
        {
            var upload = new UploadProperties
            {
                BucketName = "bucket",
                SubDirectoryInBucket = "sub",
                FileName = "file.docx",
                LocalFilePath = "/tmp/file.docx",
                AwsAccessKeyId = "key",
                AwsSecretAccessKey = "secret",
                Region = "us-east-1",
                ServiceUrl = "https://s3.example.com",
                CreatedBy = "tester",
                InputSummary = "summary",
                InputDetails = "details",
                EnableDirectDownload = true
            };

            var wordModel = new WordModel
            {
                LocalPath = "/tmp/template.docx",
                TemplatePath = new Uri("https://example.com/template.docx"),
                UploadProperties = upload,
                MinioAttachmentData = new[] { new AttachmentsData() },
                ContentControls = new List<WordContentControl>(),
                FormattingSettings = new FormattingSettings
                {
                    ProcessVoidList = true,
                    TrimAdditionalSpacingInDescriptions = true,
                    TrimAdditionalSpacingInTables = true
                },
                JsonDataList = new List<JsonData>()
            };

            var excelModel = new ExcelModel
            {
                LocalPath = "/tmp/report.xlsx",
                TemplatePath = new Uri("https://example.com/report.xlsx"),
                UploadProperties = upload,
                MinioAttachmentData = new[] { new AttachmentsData() },
                ContentControls = new List<TestReporterContentControl>(),
                JsonDataList = new List<JsonData>()
            };

            var attachments = new AttachmentsData
            {
                attachmentMinioPath = new Uri("https://example.com/file"),
                minioFileName = "file"
            };

            var jsonData = new JsonData
            {
                JsonName = "data.json",
                JsonPath = new Uri("https://example.com/data.json")
            };

            var result = new AWSUploadResult<string>
            {
                Status = true,
                StatusCode = 201,
                Data = "ok"
            };

            Assert.Equal("bucket", wordModel.UploadProperties.BucketName);
            Assert.Equal("sub", wordModel.UploadProperties.SubDirectoryInBucket);
            Assert.Equal("key", wordModel.UploadProperties.AwsAccessKeyId);
            Assert.Equal("secret", wordModel.UploadProperties.AwsSecretAccessKey);
            Assert.Equal("us-east-1", wordModel.UploadProperties.Region);
            Assert.Equal("https://s3.example.com", wordModel.UploadProperties.ServiceUrl);
            Assert.Equal("tester", wordModel.UploadProperties.CreatedBy);
            Assert.Equal("summary", wordModel.UploadProperties.InputSummary);
            Assert.Equal("details", wordModel.UploadProperties.InputDetails);
            Assert.True(wordModel.FormattingSettings.ProcessVoidList);
            Assert.True(wordModel.FormattingSettings.TrimAdditionalSpacingInDescriptions);
            Assert.True(wordModel.FormattingSettings.TrimAdditionalSpacingInTables);
            Assert.Equal("/tmp/report.xlsx", excelModel.LocalPath);
            Assert.Equal("file.docx", excelModel.UploadProperties.FileName);
            Assert.Equal("file", attachments.minioFileName);
            Assert.Equal("data.json", jsonData.JsonName);
            Assert.True(result.Status);
        }
    }
}
