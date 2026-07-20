using JsonToWord.Converters;
using JsonToWord.Models;
using JsonToWord.Models.S3;
using JsonToWord.Services.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
            var attachmentPaths = new List<String>();
            var cleanupPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var settings = new JsonSerializerSettings();
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

                            // The downloaded blob never carries ForceClean (it's assembled from the
                            // content-control response alone) — propagate it from the JsonDataList
                            // wrapper, which is the only place the original request's flag survives.
                            foreach (var cc in contentControls)
                            {
                                cc.ForceClean = cc.ForceClean || jsonData.ForceClean;
                            }

                            // Add all content controls to the wordModel
                            wordModel.ContentControls.AddRange(contentControls);
                        }
                        _aWSS3Service.CleanUp(contentControlPath);
                    }
                }
                string fullpath = ResolveTemplatePath(wordModel);
                AddCleanupPath(cleanupPaths, fullpath);
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
                AddCleanupPath(cleanupPaths, documentPath);
                _logger.LogInformation("Created word document");

                wordModel.UploadProperties.LocalFilePath = documentPath;

                if (wordModel.UploadProperties.EnableDirectDownload)
                {
                    var downloadableFile = _wordService.CreateDownloadableFile(documentPath);
                    return Ok(downloadableFile);
                }
                else
                {
                    AWSUploadResult<string> Response = await _aWSS3Service.UploadFileToMinioBucketAsync(wordModel.UploadProperties);
                    if (Response.Status)
                    {
                        return Ok(Response.Data);
                    }
                    else
                    {
                        return StatusCode(Response.StatusCode);
                    }
                }
            }
            catch (Exception e)
            {
                string logPath = @"c:\logs\prod\JsonToWord.log";
                System.IO.File.AppendAllText(logPath, string.Format("\n{0} - {1}", DateTime.Now, e));
                _logger.LogError(e, $"Error occurred while trying to create a document: {e.Message}");
                _logger.LogError($"Error Stack:\n{e.StackTrace}");
                var errorResponse = new
                {
                    message = $"Error occurred while trying to create a document: {e.Message}",
                    error = e.Message,
                    innerError = e.InnerException?.Message,
                };

                return BadRequest(JsonConvert.SerializeObject(errorResponse));
            }
            finally
            {
                foreach (var item in attachmentPaths)
                {
                    SafeCleanUp(item);
                }

                foreach (var path in cleanupPaths)
                {
                    SafeCleanUp(path);
                }
            }
        }

        private string ResolveTemplatePath(WordModel wordModel)
        {
            if (HasUsableTemplatePath(wordModel?.TemplatePath))
            {
                return _aWSS3Service.DownloadFileFromS3BucketAsync(
                    wordModel.TemplatePath,
                    wordModel?.UploadProperties?.FileName ?? "template.docx"
                );
            }

            _logger.LogInformation("TemplatePath is missing or not absolute. Building a temporary template from content controls.");
            return BuildTemporaryTemplate(wordModel);
        }

        private static bool HasUsableTemplatePath(Uri templatePath)
        {
            if (templatePath == null) return false;
            var rawValue = templatePath.ToString();
            if (string.IsNullOrWhiteSpace(rawValue)) return false;
            if (string.Equals(rawValue.Trim(), "template path", StringComparison.OrdinalIgnoreCase)) return false;
            return templatePath.IsAbsoluteUri;
        }

        private static string NormalizeFileToken(string value)
        {
            var safe = string.IsNullOrWhiteSpace(value) ? "generated-report" : value.Trim();
            var invalidChars = Path.GetInvalidFileNameChars();
            foreach (var invalid in invalidChars)
            {
                safe = safe.Replace(invalid, '-');
            }
            while (safe.Contains("--"))
            {
                safe = safe.Replace("--", "-");
            }
            safe = safe.Trim('-');
            return string.IsNullOrWhiteSpace(safe) ? "generated-report" : safe;
        }

        private static IEnumerable<string> ResolveContentControlTitles(WordModel wordModel)
        {
            var controls = wordModel?.ContentControls ?? new List<WordContentControl>();
            return controls
                .Select(control => control?.Title?.Trim())
                .Where(title => !string.IsNullOrWhiteSpace(title))
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private string BuildTemporaryTemplate(WordModel wordModel)
        {
            var fileNameBase = Path.GetFileNameWithoutExtension(wordModel?.UploadProperties?.FileName ?? "generated-report.docx");
            var safeBaseName = NormalizeFileToken(fileNameBase);
            var tempDirectoryPath = Path.Combine(Path.GetTempPath(), $"json-to-word-{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDirectoryPath);
            var tempTemplatePath = Path.Combine(tempDirectoryPath, $"{safeBaseName}.docx");
            var contentControlTitles = ResolveContentControlTitles(wordModel).ToList();

            using (var document = WordprocessingDocument.Create(tempTemplatePath, WordprocessingDocumentType.Document))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body;

                if (contentControlTitles.Count == 0)
                {
                    body.AppendChild(new Paragraph(new Run(new Text("Generated document"))));
                }
                else
                {
                    foreach (var title in contentControlTitles)
                    {
                        var sdtBlock = new SdtBlock(
                            new SdtProperties(
                                new SdtAlias { Val = title },
                                new Tag { Val = title }
                            ),
                            new SdtContentBlock(
                                new Paragraph(
                                    new Run(
                                        new Text("Click or tap here to enter text.")
                                    )
                                )
                            )
                        );
                        body.AppendChild(sdtBlock);
                        body.AppendChild(new Paragraph());
                    }
                }

                mainPart.Document.Save();
            }

            return tempTemplatePath;
        }

        private static void AddCleanupPath(ISet<string> cleanupPaths, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            cleanupPaths.Add(path);
        }

        private static bool IsGeneratedTempDirectory(string directoryPath)
        {
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                return false;
            }

            var directoryName = Path.GetFileName(
                directoryPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            );
            return directoryName.StartsWith("json-to-word-", StringComparison.OrdinalIgnoreCase);
        }

        private void SafeCleanUp(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                _aWSS3Service.CleanUp(path);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed cleaning temporary file {Path}", path);
            }

            try
            {
                var parentDirectory = Path.GetDirectoryName(path);
                if (!IsGeneratedTempDirectory(parentDirectory))
                {
                    return;
                }

                if (Directory.Exists(parentDirectory) && !Directory.EnumerateFileSystemEntries(parentDirectory).Any())
                {
                    Directory.Delete(parentDirectory, false);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Failed deleting temporary directory for {Path}", path);
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

                var document = _wordService.Create(wordModel);
                _logger.LogInformation("Created word document");

                return Ok(document);
            }
            catch (Exception e)
            {
                _logger.LogError($"Error: {e.Message}",e);
                return BadRequest($"Error: {e.Message}");
            }

        }
    }
}
