using JsonToWord.Converters;
using JsonToWord.Models;
using JsonToWord.Models.S3;
using JsonToWord.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

namespace JsonToWord.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WordController : ControllerBase
    {
        private readonly IAWSS3Service _aWSS3Service;
        private readonly IWordService _wordService;
        private readonly ILogger<WordController> _logger;

        public WordController(IAWSS3Service aWSS3Service, IWordService wordService, ILogger<WordController> logger)
        {
            _aWSS3Service = aWSS3Service;
            _wordService = wordService;
            _logger = logger;
        }

        [HttpGet("status")]
        public IActionResult GetStatus()
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);

            return Ok($"{DateTime.Now} Online - Version {versionInfo.FileVersion}");
        }

        [HttpPost("create")]
        public async Task<IActionResult> CreateWordDocument(dynamic json)
        {
            try
            {
                var settings = new JsonSerializerSettings();
                var attachmentPaths = new List<String>();
                settings.Converters.Add(new WordObjectConverter());
                WordModel wordModel = JsonConvert.DeserializeObject<WordModel>(json.ToString(), settings);
                if (wordModel.JsonDataList != null)
                {
                    wordModel.ContentControls = new List<WordContentControl>();
                    foreach (var jsonData in wordModel.JsonDataList)
                    {
                        var contentControlPath = _aWSS3Service.DownloadFileFromS3BucketAsync(jsonData.JsonPath, jsonData.JsonName);
                        using (StreamReader reader = new StreamReader(contentControlPath))
                        {
                            string contentControlJson = reader.ReadToEnd();
                            List<WordContentControl> contentControls = new List<WordContentControl>();
                            // Check if the JSON represents a list or a single object
                            if (contentControlJson.TrimStart().StartsWith("["))
                            {
                                // JSON is a list; parse it as a JArray
                                var jsonArray = JArray.Parse(contentControlJson);

                                foreach (var jsonItem in jsonArray)
                                {
                                    // Deserialize each object separately
                                    var contentControl = JsonConvert.DeserializeObject<WordContentControl>(
                                        jsonItem.ToString(),
                                        settings
                                    );
                                    contentControls.Add(contentControl);
                                }
                            }
                            else
                            {
                                // Deserialize as a single object
                                var singleContentControl = JsonConvert.DeserializeObject<WordContentControl>(contentControlJson, settings);
                                contentControls.Add(singleContentControl);
                            }

                            // Add all content controls to the wordModel
                            wordModel.ContentControls.AddRange(contentControls);
                        }
                        _aWSS3Service.CleanUp(contentControlPath);
                    }
                }
                string fullpath = _aWSS3Service.DownloadFileFromS3BucketAsync(wordModel.TemplatePath, wordModel.UploadProperties.FileName);
                wordModel.LocalPath = fullpath;
                _logger.LogInformation("Initilized word model object");
                if (wordModel.MinioAttachmentData != null)
                {
                    foreach (var item in wordModel.MinioAttachmentData)
                    {
                        attachmentPaths.Add(_aWSS3Service.DownloadFileFromS3BucketAsync(item.attachmentMinioPath, item.minioFileName));
                    }
                }
                var documentPath = _wordService.Create(wordModel);
                _logger.LogInformation("Created word document");

                _aWSS3Service.CleanUp(fullpath);

                wordModel.UploadProperties.LocalFilePath = documentPath;

                AWSUploadResult<string> Response = await _aWSS3Service.UploadFileToMinioBucketAsync(wordModel.UploadProperties);

                _aWSS3Service.CleanUp(documentPath);

                foreach (var item in attachmentPaths)
                {
                    _aWSS3Service.CleanUp(item);
                }

                if (Response.Status)
                {
                    return Ok(Response.Data);
                }
                else
                {
                    return StatusCode(Response.StatusCode);
                }

            }
            catch (Exception e)
            {
                string logPath = @"c:\logs\prod\JsonToWord.log";
                System.IO.File.AppendAllText(logPath, string.Format("\n{0} - {1}", DateTime.Now, e));
                _logger.LogError($"Error occurred while trying to create a document: {e.Message}", e);
                var errorResponse = new
                {
                    message = $"Error occurred while trying to create a document: {e.Message}",
                    error = e.Message,
                    innerError = e.InnerException?.Message,
                };

                return BadRequest(JsonConvert.SerializeObject(errorResponse));
            }
        }

        [HttpPost("create-by-file")]
        public IActionResult CreateWordDocumentByFile(dynamic json)
        {
            try
            {
                string file = json.jsonFilePath;
                string text = System.IO.File.ReadAllText(file);
                json = JObject.Parse(text);


                string test = json.ToString();
                test = test.Replace("\\\\", "\\");
                var settings = new JsonSerializerSettings();
                settings.Converters.Add(new WordObjectConverter());
                var wordModel = JsonConvert.DeserializeObject<WordModel>(json.ToString(), settings);

                _logger.LogInformation("Initilized word model object");

                var wordService = _wordService;


                var document = wordService.Create(wordModel);
                _logger.LogInformation("Created word document");

                return Ok(document);
            }
            catch (Exception e)
            {
                _logger.LogError($"Error: {e.Message}",e);
                return null;
            }

        }
    }
}