using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.EventHandlers;
using JsonToWord.Models;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
using System;
using System.IO;

public class FileService : IFileService
{
    #region Consts
    private const string IconDirectory = "Resources/Icons/";
    private const string AttachmentsFolder = "attachments";
    #endregion

    #region Event Handlers
    public event NonOfficeAttachmentEventHandler nonOfficeAttachmentEventHandler;
    #endregion

    #region Interface implementaions

    public void Insert(WordprocessingDocument document, string contentControlTitle, WordAttachment wordAttachment)
    {
        var attachedFileParagraph = AttachFileToParagraph(document.MainDocumentPart, wordAttachment);

        var sdtContentBlock = new SdtContentBlock();
        sdtContentBlock.AppendChild(attachedFileParagraph);

        var contentControlService = new ContentControlService();
        var sdtBlock = contentControlService.FindContentControl(document, contentControlTitle);
        sdtBlock.AppendChild(sdtContentBlock);
    }

    public Paragraph AttachFileToParagraph(MainDocumentPart mainPart, WordAttachment wordAttachment)
    {
        try
        {
            var fileContentType = GetFileContentType(wordAttachment.Path);
            var imageId = "";
            var iconDrawing = CreateIconImageDrawing(mainPart, wordAttachment, out imageId);

            if (fileContentType == "application/octet-stream")
            {
                TriggerNonOfficeFile();
                return AddHyperLinkNonOfficeFileParagraph(mainPart, wordAttachment, iconDrawing);
            }
            else
            {
                return CreateEmbeddedOfficeFileParagraph(mainPart, wordAttachment, imageId, fileContentType);
            }
        }
        catch (Exception ex)
        {
            string logPath = @"c:\logs\prod\JsonToWord.log";
            System.IO.File.AppendAllText(logPath, string.Format("\n{0} - {1}", DateTime.Now, ex));
            throw;
        }

    }
    #endregion

    #region Private Methods

    private Paragraph CreateEmbeddedOfficeFileParagraph(MainDocumentPart mainPart, WordAttachment wordAttachment,
     string imageId, string fileContentType)
    {
        // Convert the file to a base64 string
        var binaryData = Convert.ToBase64String(File.ReadAllBytes(wordAttachment.Path));

        if (string.IsNullOrEmpty(binaryData))
        {
            Console.WriteLine("Binary data is empty");
            return null;
        }

        using (MemoryStream dataStream = new MemoryStream(Convert.FromBase64String(binaryData)))
        {
            // Add the embedded file part to the document
            var embeddedPackagePart = mainPart.AddNewPart<EmbeddedPackagePart>(fileContentType);
            embeddedPackagePart.FeedData(dataStream);
            string embeddedPartId = mainPart.GetIdOfPart(embeddedPackagePart);

            var shapeId = GenerateShapeId();
            var prodId = GetProdId(Path.GetExtension(wordAttachment.Path));

            // Create the OLEObject element
            var oleObject = new OleObject
            {
                Type = OleValues.Embed,
                ProgId = prodId,
                DrawAspect = OleDrawAspectValues.Icon, // Display as an icon
                ObjectId = GenerateValidObjectId(),
                ShapeId = shapeId, // Unique ShapeID
                Id = embeddedPartId,
            };

            // Create ImageData element for the VML Shape
            var imageData = new DocumentFormat.OpenXml.Vml.ImageData
            {
                RelationshipId = imageId,
                Title = "" // Optional: Add title
            };

            // Create VML Shape to represent the icon
            var shape = new DocumentFormat.OpenXml.Vml.Shape
            {
                Id = shapeId, // Shape ID matching the OLEObject's ShapeID
                Type = "#_x0000_t75", // Type for embedded object representation
                Style = "width:32pt;height:32pt", // Adjust the size to fit the icon
                Ole = TrueFalseBlankValue.ToBoolean(true) // Mark as an OLE object
            };
            shape.Append(imageData); // Add the ImageData to the Shape

            // Create the EmbeddedObject (<w:object>) to hold everything
            var embeddedObject = new EmbeddedObject();
            embeddedObject.Append(shape);       // Append the VML Shape
            embeddedObject.Append(oleObject);   // Append the OLEObject

            // Create a Run to contain the EmbeddedObject
            var run = new Run();
            run.Append(embeddedObject); // Embed the object

            // Create a Run for the text with a line break
            var textRun = new Run();
            textRun.Append(new Break()); // Insert a line break between image and text
            textRun.Append(new RunProperties(new FontSize { Val = "16" })); // Font size 9 (in half-points)
            textRun.Append(new Text(wordAttachment.Name)); // Add the name of the attachment

            // Create the Paragraph to hold the Runs
            var paragraph = new Paragraph(new ParagraphProperties(
                new Justification { Val = JustificationValues.Center } // Center-align the paragraph
            ));
            paragraph.Append(run); // Add the run with the image
            paragraph.Append(textRun); // Add the run with the text

            return paragraph;
        }
    }

    private Paragraph AddHyperLinkNonOfficeFileParagraph(MainDocumentPart mainPart, WordAttachment wordAttachment, Drawing iconDrawing)
    {
        var relativePath = CopyAttachment(wordAttachment).Replace("\\", "/");

        // Create a hyperlink relationship with a relative path to the file in the 'attachments' folder
        HyperlinkRelationship hyperlinkRelationship = mainPart.AddHyperlinkRelationship(new Uri(relativePath, UriKind.Relative), true);

        // Create the hyperlink run (for text only)
        var hyperlinkRun = new Run(
            new RunProperties(
                new Underline() { Val = UnderlineValues.Single }, // Style the hyperlink text
                new Color() { Val = "0000FF" } // Hyperlink blue color
            ),
            new Text(wordAttachment.Name)
        );

        // Create a hyperlink element that wraps the hyperlink run
        var hyperlink = new Hyperlink(hyperlinkRun)
        {
            Id = hyperlinkRelationship.Id // Use the relationship ID
        };

        // Add the image and hyperlink to the document in a single paragraph
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center } // Center align the entire paragraph
            )
        );
        paragraph.Append(new Run(iconDrawing));  // Add the icon (image)
        paragraph.Append(new Break()); // Line break between image and text
        paragraph.Append(hyperlink); // Add the hyperlink for the file name


        return paragraph;
    }

    private Drawing CreateIconImageDrawing(MainDocumentPart mainPart, WordAttachment wordAttachment, out string imageId)
    {
        // Determine the icon path based on the file type
        string iconPath = GetIconPathForFileType(wordAttachment.Path);

        // Add the icon image to the document
        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png); // Assuming icons are in PNG format
        using (FileStream stream = new FileStream(iconPath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        // Get the image ID
        imageId = mainPart.GetIdOfPart(imagePart);

        // Define the size for the icon (32x32 pixels converted to EMUs - 1 pixel = 9525 EMUs)
        long widthEmu = 32 * 9525;
        long heightEmu = 32 * 9525;

        // Create the drawing object for the icon
        var drawing = new Drawing(
            new Inline(
                new Extent() { Cx = widthEmu, Cy = heightEmu }, // Set size for the image
                new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DocProperties() { Id = (UInt32Value)1U, Name = "Icon" },
                new NonVisualGraphicFrameDrawingProperties(new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Icon" },
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                new DocumentFormat.OpenXml.Drawing.Blip() { Embed = imageId },
                                new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.Transform2D(
                                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = widthEmu, Cy = heightEmu }),
                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                    new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                )
                                { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
        );

        return drawing;
    }

    private string CopyAttachment(WordAttachment wordAttachment)
    {
        var sourcePath = wordAttachment.Path;
        if (!Directory.Exists(AttachmentsFolder))
        {
            Directory.CreateDirectory(AttachmentsFolder);
        }
        string destination = Path.Combine(AttachmentsFolder, Path.GetFileName(wordAttachment.Path));
        File.Copy(sourcePath, destination, true);
        return destination;
    }

    #endregion


    #region Helper Methods

    private string GenerateShapeId()
    {
        // Use the first 6 digits of a GUID to stay within the range
        string guidPart = Guid.NewGuid().ToString("N").Substring(0, 6);
        return $"_x0000_i{guidPart}";
    }

    private string GenerateValidObjectId()
    {
        // Use the last 7 digits of a GUID to stay within the range
        string guidPart = Guid.NewGuid().ToString("N").Substring(0, 7);
        return $"_{1217000000 + Math.Abs(int.Parse(guidPart, System.Globalization.NumberStyles.HexNumber) % 483647)}";
    }

   
    private string GetFileContentType(string path)
    {
        var ext = Path.GetExtension(path);
        return ext.ToLower() switch
        {
            // Microsoft Word file formats
            ".doc" or ".dot" => "application/msword",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".dotx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
            ".docm" => "application/vnd.ms-word.document.macroEnabled.12",
            ".dotm" => "application/vnd.ms-word.template.macroEnabled.12",

            // Microsoft Excel file formats
            ".xls" => "application/vnd.ms-excel",
            ".xlsx" or ".xls" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",

            // Microsoft PowerPoint file formats
            ".ppt" => "application/vnd.ms-powerpoint",
            ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".pptm" => "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
            ".potx" => "application/vnd.openxmlformats-officedocument.presentationml.template",
            ".pot" => "application/vnd.ms-powerpoint",

            // Default content type for unknown formats
            _ => "application/octet-stream",
        };
    }

    

    private string GetProdId(string fileExt)
    {
        return fileExt.ToLower() switch
        {
            // Microsoft Word file formats
            ".doc" or ".docx" or ".docm" => "Word.Document",

            // Microsoft Word templates file formats
            ".dot" or ".dotx" or ".dotm" => "Word.Template",

            // Microsoft Excel file formats
            ".xls" or ".xlsx" or ".xlsm" => "Excel.Sheet",

            // Microsoft PowerPoint file formats
            ".ppt" or ".pptx" or ".pptm" => "PowerPoint.Show",

            // Microsoft PowerPoint file formats
            ".potx" or ".pot" => "PowerPoint.Template",

            _ => "Package",
        };
      
    }

  

    private string GetIconPathForFileType(string fileName)
    {
        string extension = Path.GetExtension(fileName).ToLower();

        // Map the file extensions to the appropriate icon paths
        switch (extension)
        {
            case ".txt":
                return $"{IconDirectory}/txt.png";
            case ".pdf":
                return $"{IconDirectory}/pdf.png";
            case ".docx":
                return $"{IconDirectory}/word.png";
            case ".xlsx":
                return $"{IconDirectory}/excel.png";
            case ".pptx":
                return $"{IconDirectory}/powerpoint.png";
            case ".csv":
                return $"{IconDirectory}/csv.png";
            case ".jpg":
            case ".jpeg":
            case ".png":
                return $"{IconDirectory}/picture.png";
            case ".mp3":
            case ".wav":
                return $"{IconDirectory}/media.png";
            default:
                return $"{IconDirectory}/default.png"; // Default icon for unknown file types
        }
    }

    #endregion

    #region Event Methods 

    private void TriggerNonOfficeFile()
    {
        nonOfficeAttachmentEventHandler?.Invoke();
    }

    #endregion
}