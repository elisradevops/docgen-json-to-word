using Amazon.S3;
using Amazon.S3.Transfer;
using JsonToWord.Models.S3;
using JsonToWord.Services.Interfaces;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Net;
using System.Threading.Tasks;
using Amazon;
using Amazon.S3.Util;
using Amazon.S3.Model;
using Microsoft.Extensions.Logging;
using System.Net.Http;
using System.Text;

namespace JsonToWord.Services
{
    public interface ITransferUtilityAdapter
    {
        Task UploadAsync(TransferUtilityUploadRequest request);
    }

    internal sealed class TransferUtilityAdapter : ITransferUtilityAdapter
    {
        private readonly TransferUtility _utility;

        public TransferUtilityAdapter(IAmazonS3 client)
        {
            _utility = new TransferUtility(client);
        }

        public Task UploadAsync(TransferUtilityUploadRequest request)
        {
            return _utility.UploadAsync(request);
        }
    }

    public class AWSS3Service : IAWSS3Service
    {
        // One client for all downloads: a client per call exhausts sockets under load. Pooled
        // connections are recycled so a storage container that restarts on a new address is picked up.
        private static readonly HttpClient SharedHttpClient = new HttpClient(
            new SocketsHttpHandler { PooledConnectionLifetime = TimeSpan.FromMinutes(5) });
        private static readonly TimeSpan DownloadTimeout = TimeSpan.FromMinutes(5);
        private const string TempDirectoryPrefix = "json-to-word-";
        // Content-control names a run's attachment directory "run-<runId>" (see its DownloadManager).
        private const string RunDirectoryPrefix = "run-";
        private const string EncodedMetadataPrefix = "utf8''";
        private const int MaxMetadataHeaderValueLength = 1024;
        private readonly ILogger<AWSS3Service> _logger;
        private readonly string localPath;
        private readonly string AwsS3BaseUrl;
        public AWSS3Service(ILogger<AWSS3Service> logger)
        {
            _logger = logger;
            localPath = "TempFiles/";
            AwsS3BaseUrl = "amazonaws.com";
        }
        // Downloads a file named by the request (template, content-control JSON). Each download gets its
        // own directory: the name comes from the request, so two concurrent requests for the same
        // document would otherwise write (and lock) the same path.
        public async Task<string> DownloadFileFromS3BucketAsync(Uri webPath, string filename)
        {
            string safeName = ToSafeFileName(filename);
            Directory.CreateDirectory(localPath);
            string requestDirectory = Path.Combine(localPath, TempDirectoryPrefix + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(requestDirectory);
            string fullPath = Path.Combine(requestDirectory, WithWebExtension(webPath, safeName));
            await DownloadToAsync(webPath, fullPath, requestDirectory);
            return fullPath;
        }

        // Downloads an attachment or picture to TempFiles/<relative path>. That path is part of the
        // contract with content-control, which writes it into the document JSON (attachmentLink) for the
        // renderer to load: content-control names it, usually "run-<runId>/<unique name>" (one directory
        // per run, so runs cannot collide or delete each other's files) and plain "<name>" when there is
        // no run id. The path comes from the request, so every segment is validated.
        public async Task<string> DownloadAttachmentAsync(Uri webPath, string filename)
        {
            string relativePath = ToSafeRelativePath(filename);
            Directory.CreateDirectory(localPath);
            string fullPath = Path.Combine(localPath, WithWebExtension(webPath, relativePath));
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
            await DownloadToAsync(webPath, fullPath, null);
            return fullPath;
        }

        // A relative path made only of plain segments: nothing rooted, no "..", nothing empty, no
        // characters a file name cannot hold. Rejected rather than trimmed, so a wrong path fails loudly.
        private static string ToSafeRelativePath(string relativePath)
        {
            var segments = (relativePath ?? string.Empty).Replace('\\', '/').Split('/');
            var invalid = Path.GetInvalidFileNameChars();
            if (segments.Length == 0 || segments.Any(segment =>
                    string.IsNullOrWhiteSpace(segment) || segment == "." || segment == ".." ||
                    segment.IndexOfAny(invalid) >= 0))
            {
                throw new ArgumentException("A valid relative file path is required for the download.", nameof(relativePath));
            }
            return string.Join("/", segments);
        }

        // The file name comes from the request: keep only its last segment, so a rooted or ".."-style
        // value can neither escape the temp directory nor, via Path.Combine, replace it.
        private static string ToSafeFileName(string filename)
        {
            string name = Path.GetFileName((filename ?? string.Empty).Replace('\\', '/'));
            if (string.IsNullOrWhiteSpace(name) || name == "." || name == "..")
            {
                throw new ArgumentException("A valid file name is required for the download.", nameof(filename));
            }
            return name;
        }

        private static string WithWebExtension(Uri webPath, string name)
        {
            return string.IsNullOrWhiteSpace(Path.GetExtension(name))
                ? name + Path.GetExtension(webPath.AbsoluteUri)
                : name;
        }

        private async Task DownloadToAsync(Uri webPath, string fullPath, string directoryToRemoveOnFailure)
        {
            try
            {
                // Bounds the whole transfer, body included (ResponseHeadersRead leaves the client's own
                // timeout covering only the headers).
                using (var cts = new CancellationTokenSource(DownloadTimeout))
                using (var response = await SharedHttpClient.GetAsync(webPath, HttpCompletionOption.ResponseHeadersRead, cts.Token))
                {
                    response.EnsureSuccessStatusCode();
                    // Streamed to disk: large templates and attachments are not buffered in memory.
                    using (var source = await response.Content.ReadAsStreamAsync())
                    using (var target = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.Read, 81920, useAsync: true))
                    {
                        await source.CopyToAsync(target, 81920, cts.Token);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Something went wrong during file download");
                try
                {
                    if (directoryToRemoveOnFailure != null)
                    {
                        // Nothing else owns this per-download directory: do not leave it behind.
                        Directory.Delete(directoryToRemoveOnFailure, true);
                    }
                    else if (File.Exists(fullPath))
                    {
                        File.Delete(fullPath);
                    }
                }
                catch (Exception cleanupError)
                {
                    _logger.LogDebug(cleanupError, "Could not remove partial download {Path}", fullPath);
                }
                throw;
            }
        }

        private bool IsOwnedTempDirectory(string directory)
        {
            var name = Path.GetFileName(directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            var parent = Path.GetDirectoryName(Path.GetFullPath(directory));
            var tempRoot = Path.GetFullPath(localPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return parent != null
                && string.Equals(parent.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), tempRoot, StringComparison.Ordinal)
                && (name.StartsWith(TempDirectoryPrefix, StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(RunDirectoryPrefix, StringComparison.OrdinalIgnoreCase));
        }

        public void CleanUp(string path)
        {
            File.Delete(path);
            // Remove the per-request or per-run directory once it is empty (best effort). Only a directory
            // of ours (known prefix) directly under TempFiles is ever removed.
            try
            {
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory)
                    && IsOwnedTempDirectory(directory)
                    && Directory.Exists(directory)
                    && !Directory.EnumerateFileSystemEntries(directory).Any())
                {
                    Directory.Delete(directory, false);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not remove temporary directory for {Path}", path);
            }
        }
        public async Task<AWSUploadResult<string>> UploadFileToS3BucketAsync(UploadProperties uploadProperties)
        {
            try
            {
                string filename = Path.GetFileName(uploadProperties.LocalFilePath);
                var transferUtilityRequest = new TransferUtilityUploadRequest()
                {
                    FilePath = uploadProperties.LocalFilePath,
                    Key = filename,
                    BucketName = uploadProperties.BucketName,
                    CannedACL = S3CannedACL.PublicReadWrite
                };
                RegionEndpoint region = RegionEndpoint.GetBySystemName(uploadProperties.Region);
                using (var util = new TransferUtility(uploadProperties.AwsAccessKeyId, uploadProperties.AwsSecretAccessKey, region))
                {
                    await util.UploadAsync(transferUtilityRequest);
                }
                var fileUrl = GenerateAwsFileUrl(uploadProperties.BucketName, filename, uploadProperties.Region);
                _logger.LogInformation("File uploaded to Amazon S3 bucket successfully");
                return fileUrl;
            }
            catch (Exception ex) when (ex is AmazonS3Exception)
            {
                _logger.LogError(ex,"Something went wrong during file upload");
                throw;
            }
        }

        public async Task<AWSUploadResult<string>> UploadFileToMinioBucketAsync(UploadProperties uploadProperties)
        {
            try
            {
                string FullBucketPath;

                if (string.IsNullOrWhiteSpace(uploadProperties.SubDirectoryInBucket))
                {
                    FullBucketPath = uploadProperties.BucketName;
                }
                else
                {
                    FullBucketPath = $"{uploadProperties.BucketName}/{uploadProperties.SubDirectoryInBucket}";
                }
                string filename = Path.GetFileName(uploadProperties.LocalFilePath);
                var transferUtilityRequest = new TransferUtilityUploadRequest()
                {
                    FilePath = uploadProperties.LocalFilePath,
                    Key = filename,
                    BucketName = FullBucketPath
                };
                
                // Add metadata including CreatedBy.
                // NOTE: `TransferUtilityUploadRequest.Metadata` expects keys WITHOUT the `x-amz-meta-` prefix.
                // The SDK will serialize them as `x-amz-meta-{key}` automatically.
                if (!string.IsNullOrEmpty(uploadProperties.CreatedBy))
                {
                    var createdBy = EncodeMetadataHeaderValue(uploadProperties.CreatedBy, 256);
                    if (!string.IsNullOrEmpty(createdBy))
                    {
                        transferUtilityRequest.Metadata.Add("createdby", createdBy);
                    }
                }
                if (!string.IsNullOrEmpty(uploadProperties.InputSummary))
                {
                    // Keep metadata safe for HTTP headers and within typical S3 metadata limits.
                    var summary = uploadProperties.InputSummary.Trim();
                    if (summary.Length > 1024)
                    {
                        summary = summary.Substring(0, 1021) + "...";
                    }
                    var encodedSummary = EncodeMetadataHeaderValue(summary, MaxMetadataHeaderValueLength);
                    if (!string.IsNullOrEmpty(encodedSummary))
                    {
                        transferUtilityRequest.Metadata.Add("inputsummary", encodedSummary);
                    }
                }
                // Store full input details as a sidecar object (avoids S3 metadata size limits).
                var hasInputDetails = !string.IsNullOrWhiteSpace(uploadProperties.InputDetails);
                // Keep the key relative to the same place as the document object (no prefixes),
                // because downstream consumers fetch from the same bucket context as the document.
                string inputDetailsObjectKey = hasInputDetails ? $"__input__/{filename}.input.json" : string.Empty;
                RegionEndpoint region = RegionEndpoint.GetBySystemName(uploadProperties.Region);
                using (var amazonClient = CreateAmazonS3Client(uploadProperties, region))
                {
                    
                    var bucketExsists = await DoesS3BucketExistAsync(amazonClient, uploadProperties.BucketName);
                    if (!bucketExsists)
                    {
                        var putBucketRequest = new PutBucketRequest
                        {
                            BucketName = uploadProperties.BucketName,
                            UseClientRegion = true
                        };
                        await amazonClient.PutBucketAsync(putBucketRequest);
                    }
                    var utility = CreateTransferUtilityAdapter(amazonClient);

                    // Best-effort: upload sidecar JSON first (so the reference is valid once the doc appears).
                    if (hasInputDetails && !string.IsNullOrWhiteSpace(inputDetailsObjectKey))
                    {
                        var inputDetailsUploaded = false;
                        try
                        {
                            var jsonBytes = Encoding.UTF8.GetBytes(uploadProperties.InputDetails);
                            using (var ms = new MemoryStream(jsonBytes))
                            {
                                var sidecarRequest = new TransferUtilityUploadRequest
                                {
                                    BucketName = FullBucketPath,
                                    Key = inputDetailsObjectKey,
                                    InputStream = ms,
                                    ContentType = "application/json",
                                    AutoCloseStream = false
                                };
                                await utility.UploadAsync(sidecarRequest);
                            }
                            inputDetailsUploaded = true;
                        }
                        catch (Exception ex)
                        {
                            // Do not fail the document upload if sidecar upload fails; just omit the reference.
                            _logger.LogWarning(ex, "Failed uploading input details sidecar for {FileName}", filename);
                        }

                        // Only attach the reference metadata if the sidecar upload succeeded.
                        if (inputDetailsUploaded)
                        {
                            var encodedInputDetailsKey = EncodeMetadataHeaderValue(inputDetailsObjectKey, MaxMetadataHeaderValueLength);
                            if (!string.IsNullOrEmpty(encodedInputDetailsKey))
                            {
                                transferUtilityRequest.Metadata.Add("inputdetailskey", encodedInputDetailsKey);
                            }
                        }
                    }
                    await utility.UploadAsync(transferUtilityRequest);
                }
                var fileUrl = GenerateMinioFileUrl(FullBucketPath, filename, uploadProperties.ServiceUrl);
                return fileUrl;
            }
            catch (Exception ex) when (ex is AmazonS3Exception)
            {
                _logger.LogError(ex, "Something went wrong during file download");
                throw;

            }
        }

        protected virtual IAmazonS3 CreateAmazonS3Client(UploadProperties uploadProperties, RegionEndpoint region)
        {
            var amazonConfig = new AmazonS3Config
            {
                AuthenticationRegion = region.SystemName,
                ServiceURL = uploadProperties.ServiceUrl,
                ForcePathStyle = true
            };
            return new AmazonS3Client(uploadProperties.AwsAccessKeyId, uploadProperties.AwsSecretAccessKey, amazonConfig);
        }

        protected virtual ITransferUtilityAdapter CreateTransferUtilityAdapter(IAmazonS3 amazonClient)
        {
            return new TransferUtilityAdapter(amazonClient);
        }

        protected virtual Task<bool> DoesS3BucketExistAsync(IAmazonS3 amazonClient, string bucketName)
        {
            return AmazonS3Util.DoesS3BucketExistV2Async(amazonClient, bucketName);
        }

        private static string EncodeMetadataHeaderValue(string value, int maxLength)
        {
            var cleaned = RemoveHeaderUnsafeCharacters(value);
            if (string.IsNullOrWhiteSpace(cleaned) || maxLength <= 0)
            {
                return string.Empty;
            }

            var normalized = cleaned.Length > maxLength ? cleaned.Substring(0, maxLength) : cleaned;
            if (IsVisibleAscii(normalized))
            {
                return normalized;
            }

            var encoded = BuildEncodedMetadataValue(normalized);
            if (encoded.Length <= maxLength)
            {
                return encoded;
            }

            var maxSourceLength = Math.Min(normalized.Length, Math.Max(1, (maxLength - EncodedMetadataPrefix.Length) / 3));
            for (var i = maxSourceLength; i > 0; i--)
            {
                encoded = BuildEncodedMetadataValue(normalized.Substring(0, i));
                if (encoded.Length <= maxLength)
                {
                    return encoded;
                }
            }

            return string.Empty;
        }

        private static string BuildEncodedMetadataValue(string value)
        {
            return string.Concat(EncodedMetadataPrefix, Uri.EscapeDataString(value));
        }

        private static string RemoveHeaderUnsafeCharacters(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var trimmed = value.Trim();
            var safeBuilder = new StringBuilder(trimmed.Length);
            foreach (var ch in trimmed)
            {
                if (ch == '\r' || ch == '\n')
                {
                    continue;
                }

                if (ch < 0x20 || ch == 0x7F)
                {
                    continue;
                }

                safeBuilder.Append(ch);
            }

            return safeBuilder.ToString();
        }

        private static bool IsVisibleAscii(string value)
        {
            foreach (var ch in value)
            {
                if (ch < 0x20 || ch > 0x7E)
                {
                    return false;
                }
            }
            return true;
        }


        public AWSUploadResult<string> GenerateAwsFileUrl(string bucketName, string key, string region, bool useRegion = true)
        {
            string publicUrl = string.Empty;
            if (useRegion)
            {
                publicUrl = $"https://{bucketName}.s3.{region}.{AwsS3BaseUrl}/{key}";
            }
            else
            {
                publicUrl = $"https://{bucketName}.s3.{AwsS3BaseUrl}/{key}";
            }
            return new AWSUploadResult<string>
            {
                Status = true,
                Data = publicUrl
            };
        }
        public AWSUploadResult<string> GenerateMinioFileUrl(string bucketName, string key, string minioServiceURL)
        {
            string publicUrl = string.Empty;
            publicUrl = $"{minioServiceURL}/{bucketName}/{key}";
            return new AWSUploadResult<string>
            {
                Status = true,
                Data = publicUrl
            };
        }
    }
}
