using Amazon.S3;
using Amazon.S3.Transfer;
using JsonToWord.Models.S3;
using JsonToWord.Services.Interfaces;
using System;
using System.IO;
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
    public class AWSS3Service : IAWSS3Service
    {
        private readonly ILogger<AWSS3Service> _logger;
        private readonly string localPath;
        private readonly string AwsS3BaseUrl;
        public AWSS3Service(ILogger<AWSS3Service> logger)
        {
            _logger = logger;
            localPath = "TempFiles/";
            AwsS3BaseUrl = "amazonaws.com";
        }
        public string DownloadFileFromS3BucketAsync(Uri webPath, string filename)
        {
            if (!Directory.Exists(localPath))
            {
                Directory.CreateDirectory(localPath);
            }
            string webExt = Path.GetExtension(webPath.AbsoluteUri);
            string fileExt = Path.GetExtension(filename);
            string fullPath;
            if (string.IsNullOrWhiteSpace(fileExt))
            {
                fullPath = localPath + filename + webExt;
            }
            else
            {
                fullPath = localPath + filename;
            }
            try
            {
                using (var client = new HttpClient())
                {
                    var response = client.GetAsync(webPath).Result;
                    response.EnsureSuccessStatusCode();
                    var fileBytes = response.Content.ReadAsByteArrayAsync().Result;
                    File.WriteAllBytes(fullPath, fileBytes);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Something went wrong during file download");
                throw;
            }
            return fullPath;
        }

        public void CleanUp(string path)
        {
            File.Delete(path);
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
                    transferUtilityRequest.Metadata.Add("createdby", uploadProperties.CreatedBy);
                }
                if (!string.IsNullOrEmpty(uploadProperties.InputSummary))
                {
                    // Keep metadata safe for HTTP headers and within typical S3 metadata limits.
                    var summary = uploadProperties.InputSummary.Trim();
                    if (summary.Length > 1024)
                    {
                        summary = summary.Substring(0, 1021) + "...";
                    }
                    transferUtilityRequest.Metadata.Add("inputsummary", summary);
                }
                // Store full input details as a sidecar object (avoids S3 metadata size limits).
                var hasInputDetails = !string.IsNullOrWhiteSpace(uploadProperties.InputDetails);
                // Keep the key relative to the same place as the document object (no prefixes),
                // because downstream consumers fetch from the same bucket context as the document.
                string inputDetailsObjectKey = hasInputDetails ? $"__input__/{filename}.input.json" : string.Empty;
                RegionEndpoint region = RegionEndpoint.GetBySystemName(uploadProperties.Region);
                var amazonConfig = new AmazonS3Config
                {
                    AuthenticationRegion = region.SystemName,
                    ServiceURL = uploadProperties.ServiceUrl,
                    ForcePathStyle = true
                };
                using (var amazonClient = new AmazonS3Client(uploadProperties.AwsAccessKeyId, uploadProperties.AwsSecretAccessKey, amazonConfig))
                {
                    
                    var bucketExsists = await AmazonS3Util.DoesS3BucketExistV2Async(amazonClient, uploadProperties.BucketName);
                    if (!bucketExsists)
                    {
                        var putBucketRequest = new PutBucketRequest
                        {
                            BucketName = uploadProperties.BucketName,
                            UseClientRegion = true
                        };
                        await amazonClient.PutBucketAsync(putBucketRequest);
                    }
                    TransferUtility utility = new TransferUtility(amazonClient);

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
                            transferUtilityRequest.Metadata.Add("inputdetailskey", inputDetailsObjectKey);
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
