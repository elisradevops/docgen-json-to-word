using System;
using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using JsonToWord.Services;
using System.IO;
using JsonToWord.Services.Interfaces;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Logging;

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
        private readonly DocumentService _documentService;
        private bool _isZipNeeded = false;
        #endregion
        
        #region Constructor
        public WordService(IContentControlService contentControlService, ITableService tableService, IPictureService pictureService, ITextService textService, IHtmlService htmlService ,IFileService fileService, ILogger<WordService> logger)
        {
            _contentControlService = contentControlService;
            _fileService = fileService;
            _htmlService = htmlService;
            _pictureService = pictureService;
            _tableService = tableService;
            _textService = textService;
            _logger = logger;
            _documentService = new DocumentService();
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
                                    _htmlService.Insert(document, contentControl.Title, (WordHtml)wordObject);
                                    break;
                                case WordObjectType.Picture:
                                    _pictureService.Insert(document, contentControl.Title, (WordAttachment)wordObject);
                                    break;
                                case WordObjectType.Paragraph:
                                    _textService.Write(document, contentControl.Title, (WordParagraph)wordObject, isUnderStandardHeading);
                                    break;
                                case WordObjectType.Table:
                                    _tableService.Insert(document, contentControl.Title, (WordTable)wordObject);
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
                
                // Clear the content control heading map after processing
                _contentControlService.ClearContentControlHeadingMap();
                
                // Save document
                document.MainDocumentPart.Document.Save();
            }

            var generatedDocPath = _isZipNeeded ? ZipDocument(documentPath) : documentPath;
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
        private string ZipDocument(string documentPath)
        {
            if (!Directory.Exists("attachments"))
            {
                throw new Exception("Attachment folder is not found");
            }

            var zipFileName = Path.ChangeExtension(documentPath, ".zip");
            CreateZipWithAttachments(zipFileName, documentPath, "attachments");

            return zipFileName;
        }

        private void CreateZipWithAttachments(string zipPath, string docxPath, string attachmentsFolder)
        {
            // Set a reasonable buffer size (e.g., 16 KB) to balance between memory usage and performance
            int bufferSize = 16 * 1024;

            // Create the ZIP file
            using (FileStream fs = File.Create(zipPath, bufferSize, FileOptions.SequentialScan))
            using (ZipOutputStream zipStream = new ZipOutputStream(fs))
            {
                // Set the compression level to a standard level compatible with Windows Explorer
                zipStream.SetLevel(6); // Compression level: 0 (no compression) to 9 (maximum compression)

                // Add the Word document to the ZIP archive
                var validPath = docxPath.Replace(":", "_");
                AddFileToZip(docxPath, Path.GetFileName(validPath), zipStream, bufferSize);

                // Add all files in the attachments folder to the ZIP archive
                foreach (string filePath in Directory.GetFiles(attachmentsFolder))
                {
                    // Add each file under the "attachments" folder in the ZIP archive
                    AddFileToZip(filePath, "attachments/" + Path.GetFileName(filePath), zipStream, bufferSize);
                }
            }
        }

        // Optimized method to add a file to the ZIP archive with Windows Explorer compatibility
        private void AddFileToZip(string filePath, string entryName, ZipOutputStream zipStream, int bufferSize)
        {
            // Create a new entry in the ZIP archive
            var entry = new ZipEntry(ZipEntry.CleanName(entryName)) // Use CleanName to ensure valid entry name
            {
                DateTime = File.GetLastWriteTime(filePath),
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