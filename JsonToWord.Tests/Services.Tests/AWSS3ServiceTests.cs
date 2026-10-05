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
                var resultPath = await service.DownloadFileFromS3BucketAsync(url, "file");

                Assert.EndsWith("file.json", resultPath);
                Assert.StartsWith("TempFiles" + Path.DirectorySeparatorChar + "json-to-word-", resultPath);
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
                var resultPath = await service.DownloadFileFromS3BucketAsync(url, "file.txt");

                Assert.EndsWith("file.txt", resultPath);
                Assert.StartsWith("TempFiles" + Path.DirectorySeparatorChar + "json-to-word-", resultPath);
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
        public async Task DownloadFileFromS3BucketAsync_SameFileNameTwice_UsesSeparatePaths_AndCleanUpRemovesDirectory()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;

            var (url1, serverTask1) = StartServer(Encoding.UTF8.GetBytes("first"), 200, "/a.docx");
            var (url2, serverTask2) = StartServer(Encoding.UTF8.GetBytes("second"), 200, "/b.docx");

            try
            {
                var first = await service.DownloadFileFromS3BucketAsync(url1, "MEWP SFTP-2026-10-05.docx");
                var second = await service.DownloadFileFromS3BucketAsync(url2, "MEWP SFTP-2026-10-05.docx");

                Assert.NotEqual(first, second);
                Assert.Equal("first", File.ReadAllText(first));
                Assert.Equal("second", File.ReadAllText(second));

                var firstDirectory = Path.GetDirectoryName(first);
                service.CleanUp(first);
                Assert.False(Directory.Exists(firstDirectory));
                Assert.True(File.Exists(second));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
                await serverTask1;
                await serverTask2;
            }
        }

        [Fact]
        public async Task DownloadAttachmentAsync_WritesToFlatTempFilesPath_ReferencedByDocumentJson()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var service = new AWSS3Service(logger.Object);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;

            var (url, serverTask) = StartServer(Encoding.UTF8.GetBytes("picture"), 200, "/img.png");

            try
            {
                var resultPath = await service.DownloadAttachmentAsync(url, "guid-1234.png");

                // content-control writes exactly "TempFiles/<name>" into attachmentLink; the renderer loads that path.
                Assert.Equal(Path.Combine("TempFiles", "guid-1234.png"), resultPath);
                Assert.Equal("picture", File.ReadAllText(Path.Combine("TempFiles", "guid-1234.png")));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
                await serverTask;
            }
        }

        [Theory]
        [InlineData("/tmp/escaped.json")]
        [InlineData("../escaped.json")]
        [InlineData("..\\escaped.json")]
        [InlineData("sub/dir/escaped.json")]
        public async Task RequestNamedDownloads_KeepOnlyTheLastSegment_InsideTheirOwnDirectory(string requestedName)
        {
            var service = new AWSS3Service(new Mock<ILogger<AWSS3Service>>().Object);
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;
            var (url, serverTask) = StartServer(Encoding.UTF8.GetBytes("a"), 200, "/a.json");

            try
            {
                var path = Path.GetFullPath(await service.DownloadFileFromS3BucketAsync(url, requestedName));
                var tempFiles = Path.GetFullPath("TempFiles") + Path.DirectorySeparatorChar;

                Assert.StartsWith(tempFiles + "json-to-word-", path);
                Assert.Equal("escaped.json", Path.GetFileName(path));
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
        public async Task DownloadAttachmentAsync_AcceptsAPerRunRelativePath_AndCleanUpRemovesTheEmptyRunDirectory()
        {
            var service = new AWSS3Service(new Mock<ILogger<AWSS3Service>>().Object);
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;
            var (url1, serverTask1) = StartServer(Encoding.UTF8.GetBytes("one"), 200, "/a.png");
            var (url2, serverTask2) = StartServer(Encoding.UTF8.GetBytes("two"), 200, "/b.png");

            try
            {
                // The same attachment name in two different runs must not collide.
                var first = await service.DownloadAttachmentAsync(url1, "run-req-1/guid.png");
                var second = await service.DownloadAttachmentAsync(url2, "run-req-2/guid.png");

                Assert.Equal(Path.Combine("TempFiles", "run-req-1", "guid.png"), first);
                Assert.Equal("one", File.ReadAllText(first));
                Assert.Equal("two", File.ReadAllText(second));

                service.CleanUp(first);
                Assert.False(Directory.Exists(Path.Combine("TempFiles", "run-req-1")));
                Assert.True(File.Exists(second));
                service.CleanUp(second);
                Assert.False(Directory.Exists(Path.Combine("TempFiles", "run-req-2")));
                Assert.True(Directory.Exists("TempFiles"));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
                await serverTask1;
                await serverTask2;
            }
        }

        [Theory]
        [InlineData("/tmp/escaped.png")]
        [InlineData("../escaped.png")]
        [InlineData("run-x/../../escaped.png")]
        [InlineData("..\\escaped.png")]
        [InlineData("run-x//guid.png")]
        [InlineData("run-x/")]
        [InlineData("run-x/bad\0name.png")]
        public async Task DownloadAttachmentAsync_RejectsPathsThatCouldLeaveTempFiles(string relativePath)
        {
            var service = new AWSS3Service(new Mock<ILogger<AWSS3Service>>().Object);
            await Assert.ThrowsAsync<ArgumentException>(() =>
                service.DownloadAttachmentAsync(new Uri("http://localhost:1/x.png"), relativePath));
        }

        [Fact]
        public void CleanUp_NeverRemovesADirectoryItDoesNotOwn()
        {
            var service = new AWSS3Service(new Mock<ILogger<AWSS3Service>>().Object);
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var originalCwd = Environment.CurrentDirectory;
            Environment.CurrentDirectory = tempDir;
            try
            {
                // An empty "run-" directory that is not directly under TempFiles, and an empty unrelated one.
                var nestedRun = Path.Combine("TempFiles", "other", "run-nested");
                var unrelated = Path.Combine("TempFiles", "keep-me");
                Directory.CreateDirectory(nestedRun);
                Directory.CreateDirectory(unrelated);
                File.WriteAllText(Path.Combine(nestedRun, "f.png"), "x");
                File.WriteAllText(Path.Combine(unrelated, "f.png"), "x");

                service.CleanUp(Path.Combine(nestedRun, "f.png"));
                service.CleanUp(Path.Combine(unrelated, "f.png"));

                Assert.True(Directory.Exists(nestedRun));
                Assert.True(Directory.Exists(unrelated));
            }
            finally
            {
                var restorePath = Directory.Exists(originalCwd) ? originalCwd : AppContext.BaseDirectory;
                Environment.CurrentDirectory = restorePath;
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task Downloads_RejectAnEmptyFileName()
        {
            var service = new AWSS3Service(new Mock<ILogger<AWSS3Service>>().Object);
            await Assert.ThrowsAsync<ArgumentException>(() => service.DownloadFileFromS3BucketAsync(new Uri("http://localhost/x.json"), ".."));
            await Assert.ThrowsAsync<ArgumentException>(() => service.DownloadAttachmentAsync(new Uri("http://localhost/x.json"), ""));
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
                await Assert.ThrowsAsync<HttpRequestException>(() => service.DownloadFileFromS3BucketAsync(url, "file"));
                Assert.Empty(Directory.GetDirectories("TempFiles"));
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
        public async Task UploadFileToMinioBucketAsync_EncodesNonAsciiMetadataAndStripsNewlines()
        {
            var logger = new Mock<ILogger<AWSS3Service>>();
            var amazonClient = new Mock<IAmazonS3>();
            var transferAdapter = new FakeTransferUtilityAdapter();
            var service = new TestableAWSS3Service(logger.Object, amazonClient.Object, transferAdapter, true);

            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            var localFile = Path.Combine(tempDir, "דוח-השוואה.docx");
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
                    CreatedBy = "ישראל כהן\n",
                    InputSummary = "Report - פע שבועי - 😀\r\n2026"
                };

                var result = await service.UploadFileToMinioBucketAsync(props);

                Assert.True(result.Status);
                Assert.Single(transferAdapter.Requests);

                var mainUpload = transferAdapter.Requests[0];
                var createdBy = mainUpload.Metadata["createdby"];
                var inputSummary = mainUpload.Metadata["inputsummary"];

                Assert.StartsWith("utf8''", createdBy);
                Assert.StartsWith("utf8''", inputSummary);
                Assert.Equal("ישראל כהן", DecodeEncodedMetadataValue(createdBy));
                Assert.Equal("Report - פע שבועי - 😀2026", DecodeEncodedMetadataValue(inputSummary));
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

        private static string DecodeEncodedMetadataValue(string value)
        {
            const string prefix = "utf8''";
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            if (!value.StartsWith(prefix, StringComparison.Ordinal))
            {
                return value;
            }

            return Uri.UnescapeDataString(value.Substring(prefix.Length));
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
