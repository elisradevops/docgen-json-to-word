using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    internal class TableService : ITableService
    {
        private IFileService _fileService;
        private IUtilsService _utilsService;
        private ILogger<TableService> _logger;

        public TableService(IFileService fileService, ILogger<TableService> logger, IUtilsService utils) {
            _fileService = fileService;
            _logger = logger;
            _utilsService = utils;
        }

        public void Insert(WordprocessingDocument document, string contentControlTitle, WordTable wordTable)
        {
            var table = CreateTable(document, wordTable);
        
            var contentControlService = new ContentControlService();
            var sdtBlock = contentControlService.FindContentControl(document, contentControlTitle);
        
            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(table);

            // Insert an empty paragraph or page Break after the table
            var emptyParagraph = new Paragraph(wordTable.InsertPageBreak
              ? (OpenXmlElement)new Run(new Break() { Type = BreakValues.Page })
              : new Run());

            sdtContentBlock.AppendChild(emptyParagraph);  // Adds an empty line
        
            sdtBlock.AppendChild(sdtContentBlock);
        
            RemoveExtraParagraphsAfterAltChunk(document);
        }

        private TableCellWidth GetTableCellWidth(string widthString, int pageWidthDxa)
        {
            if (string.IsNullOrWhiteSpace(widthString))
            {
                return new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto };
            }

            widthString = widthString.Trim().ToLowerInvariant();

            try
            {
                if (widthString.EndsWith("%"))
                {
                    double percentageWidth = _utilsService.ParseStringToDouble(widthString);
                    if (percentageWidth < 0 || percentageWidth > 100)
                    {
                        throw new ArgumentOutOfRangeException(nameof(widthString), "Percentage must be between 0 and 100.");
                    }
                    int pctWidth = (int)Math.Round(percentageWidth * 50);
                    return new TableCellWidth { Width = pctWidth.ToString(), Type = TableWidthUnitValues.Pct };
                }
                else if (widthString.EndsWith("cm"))
                {
                    double cmWidth = _utilsService.ParseStringToDouble(widthString);
                    if (cmWidth <= 0)
                    {
                        throw new ArgumentOutOfRangeException(nameof(widthString), "Width in cm must be positive.");
                    }
                    int dxaWidth = _utilsService.ConvertCmToDxa(cmWidth);
                    int pctWidth = _utilsService.ConvertDxaToPct(dxaWidth, pageWidthDxa);
                    return new TableCellWidth { Width = pctWidth.ToString(), Type = TableWidthUnitValues.Pct };
                }
                else
                {
                    throw new FormatException($"Unsupported width format: {widthString}. Use % or cm.");
                }
            }
            catch (Exception ex) when (ex is FormatException or ArgumentOutOfRangeException)
            {
                throw new ArgumentException($"Invalid width specification: {widthString}", nameof(widthString), ex);
            }
        }

        private Table CreateTable(WordprocessingDocument document, WordTable wordTable)
        {
            wordTable.RepeatHeaderRow = true;  

            var tableBorders = CreateTableBorders();
            var tableWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableLayout tableLayout = new TableLayout() { Type = TableLayoutValues.Fixed };
            var tableProperties = new TableProperties();
            tableProperties.AppendChild(tableBorders);
            tableProperties.AppendChild(tableWidth);
            tableProperties.AppendChild(tableLayout);

            int pageWidthDxa = _utilsService.GetPageWidthDxa(document.MainDocumentPart);
            

            var isHeaderRow = true;
            var table = new Table();
            table.AppendChild(tableProperties);

            var rows = wordTable.Rows;
            for (int i=0; i < rows.Count; i++)
            {
                var tableRow = new TableRow { RsidTableRowProperties = "00812C40" };

                if (wordTable.RepeatHeaderRow && isHeaderRow)
                {
                    var tableHeader = new TableHeader();

                    var tableRowProperties = new TableRowProperties();
                    tableRowProperties.AppendChild(tableHeader);

                    tableRow.AppendChild(tableRowProperties);

                    isHeaderRow = false;
                }
                var cells = rows[i].Cells; 

                for(int j=0; j < cells.Count; j++)
                {
                    var tableCellBorders = CreateTableCellBorders();

                    var tableCellWidth = GetTableCellWidth(cells[j].Width, pageWidthDxa);

                    var tableCellProperties = new TableCellProperties();
                    tableCellProperties.AppendChild(tableCellWidth);
                    tableCellProperties.AppendChild(tableCellBorders);

                    if (rows[i].MergeToOneCell)
                    {
                        var gridSpan = new GridSpan { Val = rows[i].NumberOfCellsToMerge };
                        tableCellProperties.AppendChild(gridSpan);
                    }

                    if (cells[j].Shading != null)
                    {
                        var cellShading = new Shading
                        {
                            Val = ShadingPatternValues.Clear,
                            Color = cells[j].Shading.Color,
                            Fill = cells[j].Shading.Fill,
                            ThemeFill = ThemeColorValues.Text2,
                            ThemeFillShade = cells[j].Shading.ThemeFillShade
                        };

                        tableCellProperties.AppendChild(cellShading);
                    }

                    var tableCell = new TableCell();
                    tableCell.AppendChild(tableCellProperties);

                    tableCell = AppendParagraphs(tableCell, cells[j].Paragraphs, document);

                    tableCell = AppendAttachments(tableCell, cells[j].Attachments, document);

                    
                    tableCell = AppendHtml(tableCell, cells[j].Html, document);

                    tableRow.AppendChild(tableCell);
                }

                table.AppendChild(tableRow);
            }

            return table;
        }

        private TableCell AppendHtml(TableCell tableCell, WordHtml html, WordprocessingDocument document)
        {

            if (html == null)
                return tableCell;

            if (string.IsNullOrEmpty(html.Html))
            {

                var paragraph = new Paragraph();
                tableCell.AppendChild(paragraph);

                return tableCell;
            }
            var styledHtml = WrapHtmlWithStyle(html.Html, html.Font, html.FontSize);

            var htmlService = new HtmlService();
            _logger.LogDebug("styledHtml" + styledHtml);

            var tempHtmlFile = htmlService.CreateHtmlWordDocument(styledHtml);

            var altChunkId = "altChunkId" + Guid.NewGuid().ToString("N");
            var chunk = document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            using (var fileStream = File.Open(tempHtmlFile, FileMode.Open))
            {
                chunk.FeedData(fileStream);
            }

            var altChunk = new AltChunk { Id = altChunkId };
            tableCell.AppendChild(altChunk);

            return tableCell;
        }
        private string WrapHtmlWithStyle(string originalHtml, string font, uint fontSize)
        {
            // This method wraps the HTML content with inline styles, since Word does not reliably support <style> tags in altChunk
            return $@"
                    <html>
                    <body style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>
                        {ApplyInlineStyles(originalHtml, font, fontSize)}
                    </body>
                    </html>";
        }

        // A method to apply inline styles to relevant HTML tags
        private string ApplyInlineStyles(string html, string font, uint fontSize)
        {
            // This is a basic example of how to insert inline styles for some common tags.
            // For more complex HTML, consider parsing the HTML and applying inline styles dynamically.
            return html
                .Replace("<p>", $"<p style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>")
                .Replace("<div>", $"<div style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>")
                .Replace("<span>", $"<span style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>")
                .Replace("<li>", $"<li style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>");
        }

        private TableCell AppendAttachments(TableCell tableCell, List<WordAttachment> wordAttachments, WordprocessingDocument document)
        {
            if (wordAttachments == null || !wordAttachments.Any())
                return tableCell;

            
            var pictureService = new PictureService();
            var paragraphService = new ParagraphService();

            foreach (var wordAttachment in wordAttachments)
            {
                switch (wordAttachment.Type)
                {
                    case WordObjectType.File:
                        {
                            //var embeddedFileParagraph = fileService.CreateEmbeddedObjectParagraph(document.MainDocumentPart, wordAttachment, true);
                            var embeddedFileParagraph = _fileService.AttachFileToParagraph(document.MainDocumentPart, wordAttachment);

                            if (embeddedFileParagraph != null)
                            {
                                tableCell.AppendChild(embeddedFileParagraph);
                                document.Save();
                            }
                            break;
                        }
                    case WordObjectType.Picture:
                        {
                            var drawing = pictureService.CreateDrawing(document.MainDocumentPart, wordAttachment.Path, wordAttachment.IsFlattened.GetValueOrDefault());

                            var run = new Run();
                            run.AppendChild(drawing);

                            var pictureParagraph = new Paragraph();
                            pictureParagraph.AppendChild(run);

                            // Create and add the caption below the image
                            var captionParagraph = paragraphService.CreateCaption(wordAttachment.Name);

                            tableCell.AppendChild(pictureParagraph);
                            tableCell.AppendChild(captionParagraph);
                            break;
                        }
                    default:
                        continue;
                }
            }

            return tableCell;
        }

        private TableCell AppendParagraphs(TableCell tableCell, List<WordParagraph> wordParagraphs, WordprocessingDocument document)
        {
            if (wordParagraphs == null || !wordParagraphs.Any())
                return tableCell;

            var paragraphService = new ParagraphService();

            foreach (var wordParagraph in wordParagraphs)
            {
                var paragraph = paragraphService.CreateParagraph(wordParagraph);

                if (wordParagraph.Runs != null && wordParagraph.Runs.Any())
                {
                    var runService = new RunService();

                    foreach (var wordRun in wordParagraph.Runs)
                    {
                        var run = runService.CreateRun(wordRun);
                        if (wordRun.Uri != null && wordRun.Uri != "")
                        {
                            try
                            {
                                var id = HyperlinkService.AddHyperlinkRelationship(document.MainDocumentPart, new Uri(wordRun.Uri));
                                var hyperlink = HyperlinkService.CreateHyperlink(id);
                                hyperlink.AppendChild(run);

                                paragraph.AppendChild(hyperlink);
                            }
                            catch (UriFormatException e)
                            {
                                Console.WriteLine(wordRun.Uri + " is an invalid uri \n" + e.Message);
                                paragraph.AppendChild(run);
                            }
                        }
                        else
                        {
                            paragraph.AppendChild(run);
                        }
                    }
                }

                tableCell.AppendChild(paragraph);
            }

            return tableCell;
        }

        private TableCellBorders CreateTableCellBorders()
        {
            var tableCellBorders = new TableCellBorders();
            var cellTopBorder = new TopBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var cellLeftBorder = new LeftBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var cellBottomBorder = new BottomBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var cellRightBorder = new RightBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };

            tableCellBorders.AppendChild(cellTopBorder);
            tableCellBorders.AppendChild(cellLeftBorder);
            tableCellBorders.AppendChild(cellBottomBorder);
            tableCellBorders.AppendChild(cellRightBorder);

            return tableCellBorders;
        }

        private TableBorders CreateTableBorders()
        {
            var tableBorders = new TableBorders();
            var topBorder = new TopBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var leftBorder = new LeftBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var bottomBorder = new BottomBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var rightBorder = new RightBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var insideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 };
            var insideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 };

            tableBorders.AppendChild(topBorder);
            tableBorders.AppendChild(leftBorder);
            tableBorders.AppendChild(bottomBorder);
            tableBorders.AppendChild(rightBorder);
            tableBorders.AppendChild(insideHorizontalBorder);
            tableBorders.AppendChild(insideVerticalBorder);

            return tableBorders;
        }
         private void RemoveExtraParagraphsAfterAltChunk(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart.Document.Body;
            var altChunks = body.Descendants<AltChunk>().ToList();

            foreach (var altChunk in altChunks)
            {
                // Check for a paragraph immediately following the AltChunk
                var nextParagraph = altChunk.NextSibling<Paragraph>();
                if (nextParagraph != null)
                {
                    // Check if the paragraph is empty and if so, remove it
                    if (!nextParagraph.Descendants<Run>().Any())
                    {
                        nextParagraph.Remove();
                    }
                }

                // Check for a paragraph immediately preceding the AltChunk and remove if empty
                var prevParagraph = altChunk.PreviousSibling<Paragraph>();
                if (prevParagraph != null)
                {
                    if (!prevParagraph.Descendants<Run>().Any())
                    {
                        prevParagraph.Remove();
                    }
                }
            }
        }
    }
}
