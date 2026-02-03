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
using JsonToWord.Models;

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

        private DownloadableObjectModel CreateDownloadableFile(string docPath)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(docPath);
            string fileName = Path.GetFileName(docPath);
            string base64 = Convert.ToBase64String(bytes);
            string applicationType = Path.GetExtension(docPath).ToLower() switch
            {
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                _ => "application/octet-stream"
            };

            return new DownloadableObjectModel
            {
                FileName = fileName,
                Base64 = base64,
                ApplicationType = applicationType
            };
        }
    }
}
