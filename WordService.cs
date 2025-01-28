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
        private readonly ContentControlService _contentControlService;
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
        public WordService(ITableService tableService, IPictureService pictureService, ITextService textService, IHtmlService htmlService ,IFileService fileService, ILogger<WordService> logger)
        {
            _contentControlService = new ContentControlService();
            _fileService = fileService;
            _htmlService = htmlService;
            _pictureService = pictureService;
            _tableService = tableService;
            _textService = textService;
            _documentService = new DocumentService();
            _logger = logger;
            OnSubscribeEvents();

        }
        #endregion

        #region Interface Implementations

        public string Create(WordModel _wordModel)
        {
            var documentPath = _documentService.CreateDocument(_wordModel.LocalPath);

            using (var document = WordprocessingDocument.Open(documentPath, true))
            {
                _logger.LogInformation("Starting on doc path: " + documentPath);

                foreach (var contentControl in _wordModel.ContentControls)
                {
                    _contentControlService.ClearContentControl(document, contentControl.Title, contentControl.ForceClean);

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
                                _textService.Write(document, contentControl.Title, (WordParagraph)wordObject);
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
               
            }

            var generatedDocPath = _isZipNeeded ? ZipDocument(documentPath) : documentPath;
            _logger.LogInformation("Finished on doc path: " + generatedDocPath);
            return generatedDocPath;
            //documentService.RunMacro(documentPath, "updateTableOfContent",sw);
            //log.Info("Ran Macro");
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