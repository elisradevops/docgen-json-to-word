using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.S3.Transfer;
using JsonToWord.Models.S3;
using JsonToWord.Services;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    [Collection("NonParallel")]
    public class AWSS3ServiceTests
    {
        [Fact]
        public void GenerateAwsFileUrl_UsesRegionWhenEnabled()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var result = service.GenerateAwsFileUrl("bucket", "file.docx", "us-east-1", true);

            Assert.Equal("https://bucket.s3.us-east-1.amazonaws.com/file.docx", result.Data);
            Assert.True(result.Status);
        }

        [Fact]
        public void GenerateAwsFileUrl_UsesGlobalWhenRegionDisabled()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var result = service.GenerateAwsFileUrl("bucket", "file.docx", "us-east-1", false);

            Assert.Equal("https://bucket.s3.amazonaws.com/file.docx", result.Data);
        }

        [Fact]
        public void GenerateMinioFileUrl_FormatsUrl()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var result = service.GenerateMinioFileUrl("bucket", "file.docx", "https://minio.local");

            Assert.Equal("https://minio.local/bucket/file.docx", result.Data);
        }

        [Fact]
        public void CleanUp_RemovesFile()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var filePath = Path.Combine(tempDir, "temp.txt");
            File.WriteAllText(filePath, "data");

            try
            {
                service.CleanUp(filePath);

                Assert.False(File.Exists(filePath));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task DownloadFileFromS3BucketAsync_AppendsExtension_WhenMissing()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;

            var payload = Encoding.UTF8.GetBytes("hello");
            var (url, serverTask) = StartServer(payload, 200, "/sample.json");

            try
            {
                var resultPath = service.DownloadFileFromS3BucketAsync(url, "file");

                Assert.EndsWith(Path.Combine("TempFiles", "file.json"), resultPath);
                Assert.True(File.Exists(resultPath));
                Assert.Equal("hello", File.ReadAllText(resultPath));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
                await serverTask;
            }
        }

        [Fact]
        public async Task DownloadFileFromS3BucketAsync_UsesFilenameExtension_WhenPresent()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;

            var payload = Encoding.UTF8.GetBytes("data");
            var (url, serverTask) = StartServer(payload, 200, "/sample.json");

            try
            {
                var resultPath = service.DownloadFileFromS3BucketAsync(url, "file.txt");

                Assert.EndsWith(Path.Combine("TempFiles", "file.txt"), resultPath);
                Assert.Equal("data", File.ReadAllText(resultPath));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
                await serverTask;
            }
        }

        [Fact]
        public async Task DownloadFileFromS3BucketAsync_ThrowsOnHttpError()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;

            var (url, serverTask) = StartServer(Array.Empty<byte>(), 500, "/error.json");

            try
            {
                Assert.Throws<HttpRequestException>(() => service.DownloadFileFromS3BucketAsync(url, "file"));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
                await serverTask;
            }
        }

        [Fact]
        public async Task UploadFileToMinioBucketAsync_WithMetadataAndSidecar_AddsMetadata()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var amazonClient = new Mock<IAmazonS3>();
            var transferAdapter = new FakeTransferUtilityAdapter();
            var service = new TestableAWSS3Service(logger.Object, amazonClient.Object, transferAdapter, true);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var localFile = Path.Combine(tempDir, "file.docx");
            File.WriteAllText(localFile, "content");

            try
            {
                var props = new UploadProperties
                {
                    BucketName = "bucket",
                    SubDirectoryInBucket = "sub",
                    LocalFilePath = localFile,
                    Region = "us-east-1",
                    ServiceUrl = "https://minio.local",
                    AwsAccessKeyId = "key",
                    AwsSecretAccessKey = "secret",
                    CreatedBy = "tester",
                    InputSummary = new string('a', 2000),
                    InputDetails = "{\"ok\":true}"
                };

                var result = await service.UploadFileToMinioBucketAsync(props);

                Assert.True(result.Status);
                Assert.Equal("https://minio.local/bucket/sub/file.docx", result.Data);
                Assert.Equal(2, transferAdapter.Requests.Count);

                var sidecar = transferAdapter.Requests[0];
                Assert.Equal("bucket/sub", sidecar.BucketName);
                Assert.Equal("__input__/file.docx.input.json", sidecar.Key);
                Assert.Equal("application/json", sidecar.ContentType);

                var mainUpload = transferAdapter.Requests[1];
                Assert.Equal("bucket/sub", mainUpload.BucketName);
                Assert.Equal(localFile, mainUpload.FilePath);
                Assert.Equal("tester", mainUpload.Metadata["createdby"]);
                Assert.Equal(1024, mainUpload.Metadata["inputsummary"].Length);
                Assert.EndsWith("...", mainUpload.Metadata["inputsummary"]);
                Assert.Equal("__input__/file.docx.input.json", mainUpload.Metadata["inputdetailskey"]);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task UploadFileToMinioBucketAsync_SidecarFails_OmitsInputDetailsKey()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var amazonClient = new Mock<IAmazonS3>();
            var transferAdapter = new FakeTransferUtilityAdapter();
            transferAdapter.EnqueueException(new Exception("sidecar failed"));
            var service = new TestableAWSS3Service(logger.Object, amazonClient.Object, transferAdapter, true);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var localFile = Path.Combine(tempDir, "file.docx");
            File.WriteAllText(localFile, "content");

            try
            {
                var props = new UploadProperties
                {
                    BucketName = "bucket",
                    LocalFilePath = localFile,
                    Region = "us-east-1",
                    ServiceUrl = "https://minio.local",
                    AwsAccessKeyId = "key",
                    AwsSecretAccessKey = "secret",
                    InputDetails = "{\"ok\":true}"
                };

                var result = await service.UploadFileToMinioBucketAsync(props);

                Assert.True(result.Status);
                Assert.Single(transferAdapter.Requests);
                Assert.False(transferAdapter.Requests[0].Metadata.Keys.Contains("inputdetailskey"));
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task UploadFileToMinioBucketAsync_CreatesBucketWhenMissing()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var amazonClient = new Mock<IAmazonS3>();
            amazonClient
                .Setup(c => c.PutBucketAsync(It.IsAny<PutBucketRequest>(), default))
                .ReturnsAsync(new PutBucketResponse());
            var transferAdapter = new FakeTransferUtilityAdapter();
            var service = new TestableAWSS3Service(logger.Object, amazonClient.Object, transferAdapter, false);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var localFile = Path.Combine(tempDir, "file.docx");
            File.WriteAllText(localFile, "content");

            try
            {
                var props = new UploadProperties
                {
                    BucketName = "bucket",
                    LocalFilePath = localFile,
                    Region = "us-east-1",
                    ServiceUrl = "https://minio.local",
                    AwsAccessKeyId = "key",
                    AwsSecretAccessKey = "secret"
                };

                var result = await service.UploadFileToMinioBucketAsync(props);

                Assert.True(result.Status);
                amazonClient.Verify(c => c.PutBucketAsync(It.Is<PutBucketRequest>(r => r.BucketName == "bucket"), default), Times.Once);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        private static (Uri Url, Task ServerTask) StartServer(byte[] responseBody, int statusCode, string path)
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();

            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var serverTask = Task.Run(async () =>
            {
                using var client = await listener.AcceptTcpClientAsync();
                using var stream = client.GetStream();

                var buffer = new byte[4096];
                var builder = new StringBuilder();
                while (true)
                {
                    var read = await stream.ReadAsync(buffer, 0, buffer.Length);
                    if (read <= 0)
                    {
                        break;
                    }
                    builder.Append(Encoding.ASCII.GetString(buffer, 0, read));
                    if (builder.ToString().Contains("\r\n\r\n", StringComparison.Ordinal))
                    {
                        break;
                    }
                }

                var reason = statusCode == 200 ? "OK" : "ERROR";
                var header = $"HTTP/1.1 {statusCode} {reason}\r\nContent-Length: {responseBody.Length}\r\nConnection: close\r\n\r\n";
                var headerBytes = Encoding.ASCII.GetBytes(header);
                await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
                if (responseBody.Length > 0)
                {
                    await stream.WriteAsync(responseBody, 0, responseBody.Length);
                }
                await stream.FlushAsync();
                listener.Stop();
            });

            return (new Uri($"http://127.0.0.1:{port}{path}"), serverTask);
        }

        private sealed class FakeTransferUtilityAdapter : ITransferUtilityAdapter
        {
            private readonly Queue<Exception> _exceptions = new Queue<Exception>();
            public List<TransferUtilityUploadRequest> Requests { get; } = new List<TransferUtilityUploadRequest>();

            public void EnqueueException(Exception exception)
            {
                _exceptions.Enqueue(exception);
            }

            public Task UploadAsync(TransferUtilityUploadRequest request)
            {
                if (_exceptions.Count > 0)
                {
                    throw _exceptions.Dequeue();
                }
                Requests.Add(request);
                return Task.CompletedTask;
            }
        }

        private sealed class TestableAWSS3Service : AWSS3Service
        {
            private readonly IAmazonS3 _amazonClient;
            private readonly ITransferUtilityAdapter _transferUtilityAdapter;
            private readonly bool _bucketExists;

            public TestableAWSS3Service(
                ILogger<AWSS3Service> logger,
                IAmazonS3 amazonClient,
                ITransferUtilityAdapter transferUtilityAdapter,
                bool bucketExists)
                : base(logger)
            {
                _amazonClient = amazonClient;
                _transferUtilityAdapter = transferUtilityAdapter;
                _bucketExists = bucketExists;
            }

            protected override IAmazonS3 CreateAmazonS3Client(UploadProperties uploadProperties, Amazon.RegionEndpoint region)
            {
                return _amazonClient;
            }

            protected override ITransferUtilityAdapter CreateTransferUtilityAdapter(IAmazonS3 amazonClient)
            {
                return _transferUtilityAdapter;
            }

            protected override Task<bool> DoesS3BucketExistAsync(IAmazonS3 amazonClient, string bucketName)
            {
                return Task.FromResult(_bucketExists);
            }
        }
    }
}
