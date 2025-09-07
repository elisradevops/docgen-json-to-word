using System;
using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using JsonToWord.Services;
using System.IO;
using JsonToWord.Services.Interfaces;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;

namespace JsonToWord
{
    public class WordService : IWordService, IDisposable
    {
        #region Fields
        private readonly IContentControlService _contentControlService;
        private readonly IFileService _fileService;
        private readonly ILogger<WordService> _logger;
        private readonly IPictureService _pictureService;
        private readonly ITableService _tableService;
        private readonly ITextService _textService;
        private readonly IHtmlService _htmlService;
        private readonly IVoidListService _voidListService;
        private readonly IDocumentService _documentService;
        private bool _isZipNeeded = false;
        #endregion
        
        #region Constructor
        public WordService(IContentControlService contentControlService, ITableService tableService, IPictureService pictureService, ITextService textService, IHtmlService htmlService ,IFileService fileService, IVoidListService voidListService,
         IDocumentService documentService ,ILogger<WordService> logger)
        {
            _contentControlService = contentControlService;
            _fileService = fileService;
            _htmlService = htmlService;
            _pictureService = pictureService;
            _tableService = tableService;
            _textService = textService;
            _logger = logger;
            _voidListService = voidListService;
            _documentService = documentService;
            OnSubscribeEvents();

        }
        #endregion

        #region Interface Implementations

        public string Create(WordModel _wordModel)
        {
            var documentPath = _documentService.CreateDocument(_wordModel.LocalPath);
            //If the Attachment folder already exists, delete it
            if (Directory.Exists("attachments"))
            {
                Directory.Delete("attachments", true);
            }

            using (var document = WordprocessingDocument.Open(documentPath, true))
            {
                _logger.LogInformation("Starting on doc path: " + documentPath);
                
                // PASS 1: Build the content control heading status map
                _logger.LogInformation("PASS 1: Analyzing all content controls to determine heading status");
                
                foreach (var contentControl in _wordModel.ContentControls)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(contentControl.Title))
                        {
                            _logger.LogWarning("Content control with empty title found, skipping mapping");
                            continue;
                        }
                        
                        // Find the content control
                        var sdtBlock = _contentControlService.FindContentControl(document, contentControl.Title);
                        
                        // Determine if it's under a standard heading and map the result
                        bool isUnderStandardHeading = _contentControlService.IsUnderStandardHeading(sdtBlock);
                        _contentControlService.MapContentControlHeading(contentControl.Title, isUnderStandardHeading);
                        
                        _logger.LogInformation($"Mapped content control {contentControl.Title}: Under standard heading = {isUnderStandardHeading}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"Error analyzing content control {contentControl.Title}");
                    }
                }
                
                _logger.LogInformation("PASS 2: Processing content controls with mapped heading status");
                
                // PASS 2: Process content controls using the mapped heading status
                foreach (var contentControl in _wordModel.ContentControls)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(contentControl.Title))
                        {
                            _logger.LogWarning("Content control with empty title found, skipping processing");
                            continue;
                        }
                        
                        _contentControlService.ClearContentControl(document, contentControl.Title, contentControl.ForceClean);
                        
                        // Get the content control and retrieve its heading status from the map
                        var sdtBlockCC = _contentControlService.FindContentControl(document, contentControl.Title);
                        var isUnderStandardHeading = _contentControlService.GetContentControlHeadingStatus(contentControl.Title);
                        
                        foreach (var wordObject in contentControl.WordObjects)
                        {
                            switch (wordObject.Type)
                            {
                                case WordObjectType.File:
                                    _fileService.Insert(document, contentControl.Title, (WordAttachment)wordObject);
                                    break;
                                case WordObjectType.Html:
                                    _htmlService.Insert(document, contentControl.Title, (WordHtml)wordObject, _wordModel.FormattingSettings);
                                    break;
                                case WordObjectType.Picture:
                                    _pictureService.Insert(document, contentControl.Title, (WordAttachment)wordObject);
                                    break;
                                case WordObjectType.Paragraph:
                                    _textService.Write(document, contentControl.Title, (WordParagraph)wordObject, isUnderStandardHeading);
                                    break;
                                case WordObjectType.Table:
                                    _tableService.Insert(document, contentControl.Title, (WordTable)wordObject, _wordModel.FormattingSettings);
                                    break;
                                default:
                                    throw new ArgumentOutOfRangeException();
                            }
                        }
                        document.MainDocumentPart.Document.Save();
                        _contentControlService.RemoveContentControl(document, contentControl.Title);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error processing content control: " + contentControl.Title);
                    }
                }
                
                _contentControlService.ClearContentControlHeadingMap();
                
                // Save document
                document.MainDocumentPart.Document.Save();
            }
            var voidListFiles = new List<string>();

            //Pass 3: Process Void List
            if (_wordModel.FormattingSettings.ProcessVoidList)
            {
                _logger.LogInformation("PASS 3: Processing Void List");
                voidListFiles = _voidListService.CreateVoidList(documentPath);
                _isZipNeeded = voidListFiles.Any();
            }


            var generatedDocPath = _isZipNeeded ? ZipDocument(documentPath, Directory.Exists("attachments") || voidListFiles.Any(), voidListFiles) : documentPath;
            _logger.LogInformation("Finished on doc path: " + generatedDocPath);
            return generatedDocPath;
            //documentService.RunMacro(documentPath, "updateTableOfContent",sw);
            //log.Info("Ran Macro");
        }

        public DownloadableObjectModel CreateDownloadableFile(string docPath)
        {
            // Read the document file from disk
            byte[] bytes = File.ReadAllBytes(docPath);
            
            // Get the filename without path
            string fileName = Path.GetFileName(docPath);
            
            // Base64 encode the document
            string base64 = Convert.ToBase64String(bytes);
            
            // Determine the application type based on file extension
            string applicationType = Path.GetExtension(docPath).ToLower() switch
            {
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                ".pdf" => "application/pdf",
                _ => "application/octet-stream"
            };

            return new DownloadableObjectModel
            {
                FileName = fileName,
                Base64 = base64,
                ApplicationType = applicationType
            };
        }

        public void Dispose()
        {
            OnUnsubscribeEvents();
        }
        #endregion

        #region Event Releated Methods

        private void OnNonOfficeAttachmentCaughtEvent()
        {
            if (!_isZipNeeded)
            {
                _logger.LogInformation("Non-office attachment added, the document will be zipped");
                _isZipNeeded = true;
            }
        }

        private void OnSubscribeEvents()
        {
            if(_fileService != null)
            {
                _fileService.nonOfficeAttachmentEventHandler+= OnNonOfficeAttachmentCaughtEvent;
            }
        }

        private void OnUnsubscribeEvents()
        {
            if (_fileService != null)
            {
                _fileService.nonOfficeAttachmentEventHandler -= OnNonOfficeAttachmentCaughtEvent;
            }
        }

        #endregion

        #region Zip Related Methods
        private string ZipDocument(string documentPath, bool hasAttachmentOrVoidList, List<string> voidListFilePath)
        {
            if (!hasAttachmentOrVoidList)
            {
                throw new Exception("Cannot find relevant files");
            }

            var zipFileName = Path.ChangeExtension(documentPath, ".zip");
            CreateZipWithAttachments(zipFileName, documentPath, "attachments", voidListFilePath);

            return zipFileName;
        }

        private void CreateZipWithAttachments(string zipPath, string docxPath, string attachmentsFolder, List<string> voidListFilePath)
        {
            // Set a reasonable buffer size (e.g., 16 KB) to balance between memory usage and performance
            int bufferSize = 16 * 1024;
            
            // Extract datetime from docx filename (format: ...2025-08-21-16:39:15.docx)
            DateTime? extractedDateTime = ExtractDateTimeFromFilename(docxPath);

            // Create the ZIP file
            using (FileStream fs = File.Create(zipPath, bufferSize, FileOptions.SequentialScan))
            using (ZipOutputStream zipStream = new ZipOutputStream(fs))
            {
                // Set the compression level to a standard level compatible with Windows Explorer
                zipStream.SetLevel(6); // Compression level: 0 (no compression) to 9 (maximum compression)

                // Add the Word document to the ZIP archive
                var validPath = docxPath.Replace(":", "_");
                AddFileToZip(docxPath, Path.GetFileName(validPath), zipStream, bufferSize, extractedDateTime);

                // Add the void list file if it exists
                foreach (var voidListFile in voidListFilePath)
                {
                    AddFileToZip(voidListFile, Path.GetFileName(voidListFile), zipStream, bufferSize, extractedDateTime);
                }

                // Add all files in the attachments folder to the ZIP archive if the folder exists
                if (Directory.Exists(attachmentsFolder))
                {
                    foreach (string filePath in Directory.GetFiles(attachmentsFolder))
                    {
                        // Add each file under the "attachments" folder in the ZIP archive
                        AddFileToZip(filePath, "attachments/" + Path.GetFileName(filePath), zipStream, bufferSize, extractedDateTime);
                    }
                }
            }
        }

        // Extract datetime from filename pattern: ...2025-08-21-16:39:15.docx
        private DateTime? ExtractDateTimeFromFilename(string filePath)
        {
            try
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                
                // Look for pattern: YYYY-MM-DD-HH:MM:SS at the end of filename
                var match = System.Text.RegularExpressions.Regex.Match(fileName, @"(\d{4}-\d{2}-\d{2}-\d{2}:\d{2}:\d{2})$");
                
                if (match.Success)
                {
                    string dateTimeString = match.Groups[1].Value;
                    // Convert format from YYYY-MM-DD-HH:MM:SS to YYYY-MM-DD HH:MM:SS
                    // Find the third dash and replace it with a space
                    int dashCount = 0;
                    var chars = dateTimeString.ToCharArray();
                    for (int i = 0; i < chars.Length; i++)
                    {
                        if (chars[i] == '-')
                        {
                            dashCount++;
                            if (dashCount == 3)
                            {
                                chars[i] = ' ';
                                break;
                            }
                        }
                    }
                    string formattedDateTime = new string(chars);
                    
                    if (DateTime.TryParseExact(formattedDateTime, "yyyy-MM-dd HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out DateTime result))
                    {
                        _logger.LogInformation($"Extracted datetime from filename: {result}");
                        return result;
                    }
                }
                
                _logger.LogWarning($"Could not extract datetime from filename: {fileName}");
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error extracting datetime from filename: {filePath}");
                return null;
            }
        }

        // Optimized method to add a file to the ZIP archive with Windows Explorer compatibility
        private void AddFileToZip(string filePath, string entryName, ZipOutputStream zipStream, int bufferSize, DateTime? customDateTime = null)
        {
            // Create a new entry in the ZIP archive
            var entry = new ZipEntry(ZipEntry.CleanName(entryName)) // Use CleanName to ensure valid entry name
            {
                DateTime = customDateTime ?? File.GetLastWriteTime(filePath),
                CompressionMethod = CompressionMethod.Deflated, // Use Deflate compression method
                IsUnicodeText = false, // Disable Unicode text to ensure Windows compatibility
                Size = new FileInfo(filePath).Length
            };

            // Add the entry to the ZIP stream
            zipStream.PutNextEntry(entry);

            // Write the file content to the ZIP stream with a buffered approach
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, FileOptions.SequentialScan))
            {
                byte[] buffer = new byte[bufferSize];
                int bytesRead;
                while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    zipStream.Write(buffer, 0, bytesRead);
                }
            }

            // Close the current entry in the ZIP stream
            zipStream.CloseEntry();
        }
        #endregion

    }
}