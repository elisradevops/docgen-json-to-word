using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.EventHandlers;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

public class FileService : IFileService
{
    #region Consts
    private const string IconDirectory = "Resources/Icons/";
    private const string AttachmentsFolder = "attachments";
    #endregion

    #region Fields
    private readonly IContentControlService _contentControlService;
    private readonly ILogger<FileService> _logger;
    #endregion

    #region Event Handlers
    public event NonOfficeAttachmentEventHandler nonOfficeAttachmentEventHandler;
    #endregion

    public FileService(IContentControlService contentControlService, ILogger<FileService> logger)
    {
        _contentControlService = contentControlService;
        _logger = logger;
    }


    #region Interface implementaions

    public void Insert(WordprocessingDocument document, string contentControlTitle, WordAttachment wordAttachment)
    {
        var sdtContentBlock = new SdtContentBlock();
        if (wordAttachment.IncludeAttachmentContent == false)
        {
            var attachedFileParagraph = AttachFileToParagraph(document.MainDocumentPart, wordAttachment);
            sdtContentBlock.AppendChild(attachedFileParagraph);
        }
        else
        {
            var altChunk = AddDocFileContent(document.MainDocumentPart, wordAttachment);
            if (altChunk != null)
            {
                sdtContentBlock.AppendChild(altChunk);
            }
        }

        var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);
        sdtBlock.AppendChild(sdtContentBlock);
    }

    public Paragraph AttachFileToParagraph(MainDocumentPart mainPart, WordAttachment wordAttachment)
    {
        try
        {

            if (wordAttachment == null)
            {
                throw new Exception("Word attachment is not defined");
            }

            var fileContentType = GetFileContentType(wordAttachment.Path);
            var imageId = "";
            var iconDrawing = CreateIconImageDrawing(mainPart, wordAttachment, out imageId);

            if (wordAttachment.IsLinkedFile.GetValueOrDefault() || fileContentType == "application/octet-stream")
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
            _logger.LogError($"Error occurred: {ex.Message}", ex);
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
            _logger.LogWarning($"Binary data is empty of {wordAttachment.Name}");
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
            RunProperties runProperties = new RunProperties();
            runProperties.FontSize = new FontSize { Val = "16" }; // Font size 9 (in half-points)
            textRun.RunProperties = runProperties;
            textRun.Append(new Text(wordAttachment.Name)); // Add the name of the attachment

            // Create the Paragraph to hold the Runs
            var paragraphProperties = new ParagraphProperties();
            paragraphProperties.Justification = new Justification { Val = JustificationValues.Left }; // Center align the entire paragraph
            var paragraph = new Paragraph();
            paragraph.ParagraphProperties = paragraphProperties;
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
        var runProperties = new RunProperties();
        runProperties.Underline = new Underline { Val = UnderlineValues.Single }; // Style the hyperlink text
        runProperties.Color = new Color { Val = "0000FF" }; // Hyperlink blue color

        // Create the hyperlink run (for text only)
        var hyperlinkRun = new Run();
        hyperlinkRun.RunProperties = runProperties;
        hyperlinkRun.Append(new Text(wordAttachment.Name));

        // Create a hyperlink element that wraps the hyperlink run
        var hyperlink = new Hyperlink(hyperlinkRun)
        {
            Id = hyperlinkRelationship.Id // Use the relationship ID
        };

        // Add the image and hyperlink to the document in a single paragraph
        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Justification = new Justification { Val = JustificationValues.Left }; // Center align the entire paragraph
        
        var paragraph = new Paragraph();
        paragraph.ParagraphProperties = paragraphProperties;
        paragraph.Append(new Run(iconDrawing));  // Add the icon (image)
        paragraph.Append(new Run(new Break())); // Line break between image and text
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
        var guidFileName = Path.GetFileName(wordAttachment.Path);
        var extension = Path.GetExtension(guidFileName);

        string destination = Path.Combine(AttachmentsFolder, wordAttachment.Name + extension);
        while (File.Exists(destination))
        {
            string uniqueId = Guid.NewGuid().ToString("N").Substring(0, 4);
            destination = Path.Combine(AttachmentsFolder, $"{wordAttachment.Name}-(CopyID-{uniqueId}){extension}");
        }
        File.Copy(sourcePath, destination, false);
        return destination;
    }

    private AltChunk AddDocFileContent(MainDocumentPart mainPart, WordAttachment wordAttachment)
    {
        try
        {
            var altChunkId = "altChunkId" + Guid.NewGuid().ToString("N");
            var chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            // Clean the source document and get the cleaned content as a memory stream
            using (MemoryStream cleanedDocumentStream = CleanWordDocument(wordAttachment.Path))
            {
                cleanedDocumentStream.Position = 0; // Reset the stream position
                chunk.FeedData(cleanedDocumentStream);
            }

            var altChunk = new AltChunk { Id = altChunkId };
            return altChunk;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, $"Cannot add {wordAttachment.Name} document content");
        }

        return null;
    }

    private MemoryStream CleanWordDocument(string documentPath)
    {
        // Create a memory stream to hold the cleaned document
        var cleanedStream = new MemoryStream();

        try
        {
            // Open the source document as read-only
            using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(documentPath, false))
            {
                // Clone the source document into the memory stream
                sourceDoc.Clone(cleanedStream, true);
            }

            // Open the cloned document for editing
            using (WordprocessingDocument cleanedDoc = WordprocessingDocument.Open(cleanedStream, true))
            {
                var body = cleanedDoc.MainDocumentPart.Document.Body;

                // Remove redundant elements at the beginning
                while (body.FirstChild is Paragraph paragraph && IsRedundantParagraph(paragraph))
                {
                    paragraph.Remove();
                }

                // Remove redundant elements at the end
                while (body.LastChild is Paragraph paragraph && IsRedundantParagraph(paragraph))
                {
                    paragraph.Remove();
                }

                // Remove headers and footers from the document
                var headers = cleanedDoc.MainDocumentPart.HeaderParts.ToList();
                foreach (var header in headers)
                {
                    cleanedDoc.MainDocumentPart.DeletePart(header);
                }

                var footers = cleanedDoc.MainDocumentPart.FooterParts.ToList();
                foreach (var footer in footers)
                {
                    cleanedDoc.MainDocumentPart.DeletePart(footer);
                }

                // Remove header and footer references in section properties
                foreach (var sectPr in cleanedDoc.MainDocumentPart.Document.Descendants<SectionProperties>())
                {
                    // Remove header references
                    sectPr.Elements<HeaderReference>().ToList().ForEach(h => h.Remove());

                    // Remove footer references
                    sectPr.Elements<FooterReference>().ToList().ForEach(f => f.Remove());

                    // Remove page size and margin settings to adopt target document settings
                    sectPr.RemoveAllChildren<PageSize>();
                    sectPr.RemoveAllChildren<PageMargin>();
                }

                // Convert heading styles to regular text while preserving formatting
                ConvertHeadingsToRegularText(body);

                // Remove numbering definitions to prevent conflicts
                RemoveNumberingReferences(cleanedDoc);

                // Make style IDs unique to prevent conflicts with the target document
                PrefixStyleIds(cleanedDoc);

                // Make all IDs unique using direct XML manipulation
                MakeAllIdsUnique(cleanedDoc);

                // Add a section break at the beginning to isolate formatting
                InsertSectionBreak(body);

                // Save changes to the document
                cleanedDoc.MainDocumentPart.Document.Save();
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error cleaning document");

            // If we had an error, reset the stream and create a minimal valid document
            cleanedStream.SetLength(0);
            using (WordprocessingDocument minimalDoc = WordprocessingDocument.Create(cleanedStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var mainPart = minimalDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(
                    new Run(new Text("Docgen error processing document - see log for details")))));
                mainPart.Document.Save();
            }
        }

        cleanedStream.Position = 0;
        return cleanedStream;
    }

    /// <summary>
    /// Adds a prefix to all style IDs to prevent conflicts with target document styles
    /// </summary>
    private void PrefixStyleIds(WordprocessingDocument doc)
    {
        var stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
        if (stylesPart == null) return;

        // Generate a unique prefix for this document's styles
        string stylePrefix = "s" + Guid.NewGuid().ToString("N").Substring(0, 8) + "_";

        // Build a mapping of old style IDs to new prefixed style IDs
        var styleIdMap = new Dictionary<string, string>();

        // First pass - collect all style IDs and create new prefixed versions
        foreach (var style in stylesPart.Styles.Elements<Style>())
        {
            if (style.StyleId != null)
            {
                // Skip default styles to maintain compatibility
                if (style.StyleId == "Normal" || style.StyleId == "DefaultParagraphFont" ||
                    style.StyleId == "TableNormal" || style.StyleId == "NoList")
                {
                    continue;
                }

                string oldId = style.StyleId;
                string newId = stylePrefix + oldId;
                styleIdMap[oldId] = newId;
            }
        }

        // Second pass - update style IDs and any references to other styles
        foreach (var style in stylesPart.Styles.Elements<Style>().ToList())
        {
            if (style.StyleId != null && styleIdMap.ContainsKey(style.StyleId))
            {
                style.StyleId = styleIdMap[style.StyleId];

                // Update base style references
                if (style.BasedOn != null && style.BasedOn.Val != null &&
                    styleIdMap.ContainsKey(style.BasedOn.Val))
                {
                    style.BasedOn.Val = styleIdMap[style.BasedOn.Val];
                }

                // Update next style references
                if (style.NextParagraphStyle != null && style.NextParagraphStyle.Val != null &&
                    styleIdMap.ContainsKey(style.NextParagraphStyle.Val))
                {
                    style.NextParagraphStyle.Val = styleIdMap[style.NextParagraphStyle.Val];
                }
            }
        }

        // Update references to styles in the document content
        foreach (var paragraph in doc.MainDocumentPart.Document.Descendants<Paragraph>())
        {
            if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val != null &&
                styleIdMap.ContainsKey(paragraph.ParagraphProperties.ParagraphStyleId.Val))
            {
                paragraph.ParagraphProperties.ParagraphStyleId.Val =
                    styleIdMap[paragraph.ParagraphProperties.ParagraphStyleId.Val];
            }

            foreach (var run in paragraph.Elements<Run>())
            {
                if (run.RunProperties?.RunStyle?.Val != null &&
                    styleIdMap.ContainsKey(run.RunProperties.RunStyle.Val))
                {
                    run.RunProperties.RunStyle.Val = styleIdMap[run.RunProperties.RunStyle.Val];
                }
            }
        }

        // Update references to styles in tables
        foreach (var table in doc.MainDocumentPart.Document.Descendants<Table>())
        {  
            var props = table.GetFirstChild<TableProperties>();
            if (props != null) {
                if (props.TableStyle?.Val != null &&
                  styleIdMap.ContainsKey(props.TableStyle.Val))
                {
                    props.TableStyle.Val = styleIdMap[props.TableStyle.Val];
                }
            }
          
        }
    }

    private void ConvertHeadingsToRegularText(Body body)
    {
        // Find all paragraphs with heading styles
        var headingParagraphs = body.Descendants<Paragraph>()
            .Where(p => p.ParagraphProperties?.ParagraphStyleId != null &&
                   p.ParagraphProperties.ParagraphStyleId.Val != null &&
                   p.ParagraphProperties.ParagraphStyleId.Val.Value.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var paragraph in headingParagraphs)
        {
            // Preserve the text and basic formatting, but remove the heading style
            if (paragraph.ParagraphProperties != null)
            {
                // Convert to a custom style or normal text while keeping the formatting
                var styleId = paragraph.ParagraphProperties.ParagraphStyleId;
                if (styleId != null)
                {
                    // Create a copy of formatting properties but change the style
                    styleId.Val = "Normal";

                    // Remove any numbering references specifically
                    if (paragraph.ParagraphProperties.NumberingProperties != null)
                    {
                        paragraph.ParagraphProperties.NumberingProperties.Remove();
                    }
                }
            }
        }
    }

    /// <summary>
    /// Insert a section break at the beginning of the document to isolate formatting
    /// </summary>
    private void InsertSectionBreak(Body body)
    {
        // Create section properties with a continuous section break
        var sectionProps = new SectionProperties(
            new SectionType { Val = SectionMarkValues.Continuous }
        );

        // If the body has content, add the section break to the first paragraph
        if (body.FirstChild != null)
        {
            if (body.FirstChild is Paragraph firstParagraph)
            {
                // If the paragraph doesn't have properties, add them
                if (firstParagraph.ParagraphProperties == null)
                {
                    firstParagraph.PrependChild(new ParagraphProperties());
                }

                // Add section properties to the paragraph
                firstParagraph.ParagraphProperties.SectionProperties = sectionProps;
            }
            else
            {
                // Insert a new paragraph with section break before the first element
                body.InsertBefore(
                    new Paragraph(
                        new ParagraphProperties(sectionProps)
                    ),
                    body.FirstChild
                );
            }
        }
        else
        {
            // If body is empty, just add a paragraph with section break
            body.AppendChild(
                new Paragraph(
                    new ParagraphProperties(sectionProps)
                )
            );
        }
    }

    private void RemoveNumberingReferences(WordprocessingDocument document)
    {
        // Get the numbering part (if it exists)
        var numberingPart = document.MainDocumentPart.NumberingDefinitionsPart;
        if (numberingPart != null)
        {
            // Get all paragraphs with numbering references
            var paragraphsWithNumbering = document.MainDocumentPart.Document.Descendants<Paragraph>()
                .Where(p => p.ParagraphProperties?.NumberingProperties != null)
                .ToList();

            // Remove numbering properties from these paragraphs
            foreach (var paragraph in paragraphsWithNumbering)
            {
                paragraph.ParagraphProperties.NumberingProperties.Remove();
            }

            // Optionally remove the entire numbering part
            document.MainDocumentPart.DeletePart(numberingPart);
        }
    }

    private void MakeAllIdsUnique(WordprocessingDocument doc)
    {
        // Let's use a direct XML manipulation approach to ensure all IDs are unique

        // First, save the document so all changes are written to the XML
        doc.MainDocumentPart.Document.Save();

        // Get the XML document
        XmlDocument xmlDoc = new XmlDocument();
        using (Stream xmlStream = doc.MainDocumentPart.GetStream(FileMode.Open, FileAccess.Read))
        {
            xmlDoc.Load(xmlStream);
        }

        // Create a namespace manager
        XmlNamespaceManager nsManager = new XmlNamespaceManager(xmlDoc.NameTable);
        nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        nsManager.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        nsManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        nsManager.AddNamespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        nsManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // Used to track IDs to ensure uniqueness
        var usedIds = new HashSet<string>();

        // Find ALL elements with ID attributes, regardless of namespace
        // This pattern selects any element that has an attribute named 'id' or 'Id'
        XmlNodeList allElementsWithId = xmlDoc.SelectNodes("//*[@id or @Id]", nsManager);
        if (allElementsWithId != null)
        {
            int idCounter = 1000000;

            foreach (XmlNode node in allElementsWithId)
            {
                // Check for 'id' attribute (lowercase)
                XmlAttribute idAttr = node.Attributes["id"];
                if (idAttr != null)
                {
                    string newId = (idCounter++).ToString();
                    idAttr.Value = newId;
                    usedIds.Add(newId);
                }

                // Check for 'Id' attribute (uppercase first letter)
                XmlAttribute IdAttr = node.Attributes["Id"];
                if (IdAttr != null)
                {
                    string newId = (idCounter++).ToString();
                    IdAttr.Value = newId;
                    usedIds.Add(newId);
                }
            }
        }

        // Specifically target the docPr elements which are causing the issues
        XmlNodeList docPrElements = xmlDoc.SelectNodes("//wp:docPr", nsManager);
        if (docPrElements != null)
        {
            int idCounter = 2000000;

            foreach (XmlNode node in docPrElements)
            {
                XmlAttribute idAttr = node.Attributes["id"];
                if (idAttr != null)
                {
                    string newId = (idCounter++).ToString();
                    idAttr.Value = newId;
                    usedIds.Add(newId);
                }
            }
        }

        // Save the modified XML back to the document
        using (Stream xmlStream = doc.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            xmlDoc.Save(xmlStream);
        }

    }

    private bool IsRedundantParagraph(Paragraph paragraph)
    {
        // Check if the paragraph has no runs (completely empty)
        if (!paragraph.Elements<Run>().Any())
            return true;

        // Check if all runs contain only breaks or whitespace
        return paragraph.Elements<Run>().All(run =>
            !run.Elements<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text)) &&
            !run.Elements<FieldChar>().Any() && // Exclude fields
            run.Elements<Break>().All(br => br.Type == BreakValues.TextWrapping || br.Type == BreakValues.Page));
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
        string extension = Path.GetExtension(fileName);

        return extension.ToLower() switch
        {
            // Microsoft Word file formats
            ".doc" or ".docx" or ".docm" => $"{IconDirectory}/word.png",

            // Microsoft Word templates file formats
            ".dot" or ".dotx" or ".dotm" => $"{IconDirectory}/word.png",

            // Microsoft Excel file formats
            ".xls" or ".xlsx" or ".xlsm" => $"{IconDirectory}/excel.png",

            // Microsoft PowerPoint file formats
            ".ppt" or ".pptx" or ".pptm" => $"{IconDirectory}/powerpoint.png",

            // Microsoft PowerPoint file formats
            ".potx" or ".pot" => $"{IconDirectory}/powerpoint.png",

            ".txt" => $"{IconDirectory}/txt.png",

            ".pdf" => $"{IconDirectory}/pdf.png",
            
            ".csv" => $"{IconDirectory}/csv.png",

            ".jpg" or ".jpeg" or ".png" => $"{IconDirectory}/picture.png",

            ".mp3" or ".wav" => $"{IconDirectory}/media.png",

            ".xml" => $"{IconDirectory}/xml.png",

            ".zip" or ".7z" => $"{IconDirectory}/zip.png",

            ".rar" => $"{IconDirectory}/rar.png",

            _ => $"{IconDirectory}/default.png",
        };
    }

    #endregion

    #region Event Methods 

    private void TriggerNonOfficeFile()
    {
        nonOfficeAttachmentEventHandler?.Invoke();
    }

    #endregion
}