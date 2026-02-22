using JsonToWord.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.Reflection;
using System;
using System.Threading.Tasks;
using Newtonsoft.Json;
using JsonToWord.Converters;
using JsonToWord.Models;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using JsonToWord.Models.S3;
using ICSharpCode.SharpZipLib.Zip;

namespace JsonToWord.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IAWSS3Service _aWSS3Service;
        private readonly IExcelService _excelService;
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(IAWSS3Service aWSS3Service, IExcelService excelService, ILogger<ExcelController> logger)
        {
            _aWSS3Service = aWSS3Service;
            _excelService = excelService;
            _logger = logger;
        }

        [HttpGet("status")]
        public IActionResult GetStatus()
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);

            return Ok($"{DateTime.Now} Online - Version {versionInfo.FileVersion}");
        }

        [HttpPost("create")]
        public async Task<IActionResult> CreateExcelDocument(dynamic json)
        {
            try
            {
                var settings = new JsonSerializerSettings();
                settings.Converters.Add(new TestReporterConverter());
                ExcelModel excelModel = JsonConvert.DeserializeObject<ExcelModel>(json.ToString(), settings);
                if (excelModel.JsonDataList != null)
                {
                    excelModel.ContentControls = new List<TestReporterContentControl>();
                    foreach (var jsonData in excelModel.JsonDataList)
                    {
                        var contentControlPath = _aWSS3Service.DownloadFileFromS3BucketAsync(jsonData.JsonPath, jsonData.JsonName);
                        using (StreamReader reader = new StreamReader(contentControlPath))
                        {
                            string contentControlJson = reader.ReadToEnd();
                            List<TestReporterContentControl> contentControls = new List<TestReporterContentControl>();
                            // Check if the JSON represents a list or a single object
                            if (contentControlJson.TrimStart().StartsWith("["))
                            {
                                // JSON is a list; parse it as a JArray
                                var jsonArray = JArray.Parse(contentControlJson);

                                foreach (var jsonItem in jsonArray)
                                {
                                    // Deserialize each object separately
                                    var contentControl = JsonConvert.DeserializeObject<TestReporterContentControl>(
                                        jsonItem.ToString(),
                                        settings
                                    );
                                    contentControls.Add(contentControl);
                                }
                            }
                            else
                            {
                                // Deserialize as a single object
                                var singleContentControl = JsonConvert.DeserializeObject<TestReporterContentControl>(contentControlJson, settings);
                                contentControls.Add(singleContentControl);
                            }
                            excelModel.ContentControls.AddRange(contentControls);
                        }
                        _aWSS3Service.CleanUp(contentControlPath);
                    }
                }

                // Ensure the filename has '.xlsx' extension
                if (!excelModel.UploadProperties.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    excelModel.UploadProperties.FileName += ".xlsx";
                }

                // Set the LocalPath using the updated filename
                excelModel.LocalPath = Path.Combine("TempFiles", excelModel.UploadProperties.FileName);
                _logger.LogInformation("Initilized word model object");

                var spreadsheetPath = _excelService.CreateExcelDocument(excelModel);
                _logger.LogInformation("Excel document created successfully");
                excelModel.UploadProperties.LocalFilePath = spreadsheetPath;

                if (excelModel.UploadProperties.EnableDirectDownload)
                {
                    var downloadableFile = CreateDownloadableFile(spreadsheetPath);
                    _aWSS3Service.CleanUp(spreadsheetPath);
                    return Ok(downloadableFile);
                }
                else
                {
                    AWSUploadResult<string> Response = await _aWSS3Service.UploadFileToMinioBucketAsync(excelModel.UploadProperties);

                    _aWSS3Service.CleanUp(spreadsheetPath);

                    if (Response.Status)
                    {
                        return Ok(Response.Data);
                    }
                    return StatusCode(Response.StatusCode);
                }
            }
            catch (Exception e)
            {
                string logPath = @"c:\logs\prod\JsonToWord.log";
                System.IO.File.AppendAllText(logPath, string.Format("\n{0} - {1}", DateTime.Now, e));
                _logger.LogError(e, $"Error occurred while trying to create a spreadsheet: {e.Message}");
                _logger.LogError($"Error Stack:\n{e.StackTrace}");
                var errorResponse = new
                {
                    message = $"Error occurred while trying to create a document: {e.Message}",
                    error = e.Message,
                    innerError = e.InnerException?.Message,
                };

                return BadRequest(JsonConvert.SerializeObject(errorResponse));
            }
        }

        [HttpPost("create-zip")]
        public async Task<IActionResult> CreateExcelZipPackage(dynamic json)
        {
            string zipPath = null;
            try
            {
                var zipModel = JsonConvert.DeserializeObject<ExcelZipPackageModel>(json.ToString());
                if (zipModel?.UploadProperties == null)
                {
                    return BadRequest(JsonConvert.SerializeObject(new
                    {
                        message = "Missing uploadProperties in zip payload",
                    }));
                }

                if (zipModel.Files == null || zipModel.Files.Count == 0)
                {
                    return BadRequest(JsonConvert.SerializeObject(new
                    {
                        message = "No files provided for zip package",
                    }));
                }

                var zipFileName = EnsureZipFileName(zipModel.UploadProperties.FileName);
                if (!Directory.Exists("TempFiles"))
                {
                    Directory.CreateDirectory("TempFiles");
                }
                zipPath = Path.Combine("TempFiles", zipFileName);

                using (var stream = System.IO.File.Create(zipPath))
                using (var zipStream = new ZipOutputStream(stream))
                {
                    zipStream.SetLevel(6);
                    foreach (var file in zipModel.Files)
                    {
                        if (string.IsNullOrWhiteSpace(file?.FileName) || string.IsNullOrWhiteSpace(file?.Base64))
                        {
                            continue;
                        }

                        var entryName = Path.GetFileName(file.FileName.Trim());
                        if (string.IsNullOrWhiteSpace(entryName))
                        {
                            continue;
                        }

                        var payload = Convert.FromBase64String(file.Base64.Trim());
                        var entry = new ZipEntry(entryName)
                        {
                            DateTime = DateTime.Now,
                        };
                        zipStream.PutNextEntry(entry);
                        zipStream.Write(payload, 0, payload.Length);
                        zipStream.CloseEntry();
                    }
                    zipStream.Finish();
                }

                if (zipModel.UploadProperties.EnableDirectDownload)
                {
                    var downloadableFile = CreateDownloadableFile(zipPath);
                    _aWSS3Service.CleanUp(zipPath);
                    return Ok(downloadableFile);
                }

                zipModel.UploadProperties.LocalFilePath = zipPath;
                AWSUploadResult<string> response = await _aWSS3Service.UploadFileToMinioBucketAsync(zipModel.UploadProperties);
                _aWSS3Service.CleanUp(zipPath);

                if (response.Status)
                {
                    return Ok(response.Data);
                }
                return StatusCode(response.StatusCode);
            }
            catch (Exception e)
            {
                if (!string.IsNullOrWhiteSpace(zipPath) && System.IO.File.Exists(zipPath))
                {
                    _aWSS3Service.CleanUp(zipPath);
                }

                _logger.LogError(e, $"Error occurred while trying to create zip package: {e.Message}");
                var errorResponse = new
                {
                    message = $"Error occurred while trying to create zip package: {e.Message}",
                    error = e.Message,
                    innerError = e.InnerException?.Message,
                };
                return BadRequest(JsonConvert.SerializeObject(errorResponse));
            }
        }

        private DownloadableObjectModel CreateDownloadableFile(string docPath)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(docPath);
            string fileName = Path.GetFileName(docPath);
            string base64 = Convert.ToBase64String(bytes);
            string applicationType = Path.GetExtension(docPath).ToLower() switch
            {
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".zip" => "application/zip",
                _ => "application/octet-stream"
            };

            return new DownloadableObjectModel
            {
                FileName = fileName,
                Base64 = base64,
                ApplicationType = applicationType
            };
        }

        private string EnsureZipFileName(string rawFileName)
        {
            var fileName = string.IsNullOrWhiteSpace(rawFileName) ? "report.zip" : rawFileName.Trim();
            if (!fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
            {
                fileName += ".zip";
            }
            return fileName;
        }
    }
}
