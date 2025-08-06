using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using HtmlToOpenXml.Custom;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class HtmlService : IHtmlService
    {
        private readonly IContentControlService _contentControlService;
        private readonly IDocumentValidatorService _documentValidator;
        private readonly ILogger<HtmlService> _logger;
        private readonly IPictureService _pictureService;
        private readonly IParagraphService _paragraphService;

        // Common lists – these are not exhaustive, but cover many elements.
        private readonly HashSet<string> inlineElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "a", "abbr", "acronym", "b", "bdo", "big", "br", "button", "cite", "code", "dfn",
            "em", "i", "img", "input", "kbd", "label", "map", "object", "output", "q", "samp",
            "script", "select", "small", "span", "strong", "sub", "sup", "textarea", "time", "tt", "var"
        };

        private readonly HashSet<string> blockElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "address", "article", "aside", "blockquote", "canvas", "dd", "div", "dl", "dt", "fieldset",
            "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr",
            "li", "main", "nav", "noscript", "ol", "p", "pre", "section", "table", "ul", "video"
        };

        public HtmlService(IContentControlService contentControlService,IDocumentValidatorService documentValidator, ILogger<HtmlService> logger, IPictureService pictureService, IParagraphService paragraphService)
        {
            _contentControlService = contentControlService;
            _logger = logger;
            _documentValidator = documentValidator;
            _pictureService = pictureService;
            _paragraphService = paragraphService;
        }
        public void Insert(WordprocessingDocument document, string contentControlTitle, WordHtml wordHtml, FormattingSettings formattingSettings)
        {
            var elements = ConvertHtmlToOpenXmlElements(wordHtml, document);

            // Always resize images from HTML content to fit properly in the document
            foreach (var paragraph in elements.OfType<Paragraph>())
            {
                ResizeImagesInParagraph(paragraph);
            }

            // Apply tight spacing to HTML-generated paragraphs if TrimAdditionalSpacingInTables is enabled
            if (formattingSettings?.TrimAdditionalSpacingInDescriptions == true)
            {
                foreach (var paragraph in elements.OfType<Paragraph>())
                {
                    _paragraphService.ApplyTightSpacing(paragraph);
                }
            }

            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();

            sdtContentBlock.Append(elements);

            sdtBlock.AppendChild(sdtContentBlock);
        }


        public IEnumerable<OpenXmlCompositeElement> ConvertHtmlToOpenXmlElements(WordHtml wordHtml, WordprocessingDocument document)
        {
            bool isHtmlEmpty = string.IsNullOrEmpty(wordHtml.Html);

            var errors = ValidateHtmlStructure(wordHtml.Html);

            var html = WrapHtmlWithStyle(wordHtml.Html, wordHtml.Font, wordHtml.FontSize);

            html = RemoveWordHeading(html);

            html = FixBullets(html);
            Uri baseImageUrl;
            try
            {
                baseImageUrl = new Uri(System.IO.Path.GetFullPath(Environment.CurrentDirectory));
            }
            catch (Exception)
            {
                // Fallback to app directory in Docker
                baseImageUrl = new Uri("file:///app/");
            }

            // Rest of your existing method
            var converter = new HtmlConverter(document.MainDocumentPart, new HtmlToOpenXml.Custom.IO.DefaultWebRequest()
            {
                BaseImageUrl = baseImageUrl
            });
            converter.ContinueNumbering = false;
            converter.SupportsHeadingNumbering = false;

            try
            {
                var elements = converter.Parse(html);
                if(elements.Count == 0 && errors.Count > 0 || isHtmlEmpty)
                {
                    if(errors.Count > 0)
                    {
                        string errorMessage = string.Join("\n", errors);
                        _logger.LogError("Errors found in the html: " + errorMessage);
                    }
                    throw new Exception("Invalid HTML Format");
                }
                return elements;
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
                _logger.LogError(ex, $"DocGen ran into an issue parsing the html due to: {errorMessage}");
                _logger.LogError($"The html that caused the issue is: {wordHtml.Html}");

                string errorHtml = "<html><head></head><body><p style='color: red'><b>Docgen Error: Invalid HTML Format: " + errorMessage + "</b></p></body></html>";
                var elements = converter.Parse(errorHtml);
                return elements;
            }
        }

        

        private string WrapHtmlWithStyle(string originalHtml, string font, uint fontSize)
        {
            // Check if the originalHtml is already wrapped with <html> tags
            if (originalHtml.TrimStart().StartsWith("<html>", StringComparison.OrdinalIgnoreCase) &&
                originalHtml.TrimEnd().EndsWith("</html>", StringComparison.OrdinalIgnoreCase))
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(originalHtml);

                var bodyNode = doc.DocumentNode.SelectSingleNode("//body");

                if (bodyNode != null)
                {
                    string newStyle = $"font-family: {font}, sans-serif; font-size: {fontSize}pt;";

                    string existingStyle = bodyNode.GetAttributeValue("style", "");
                    string combinedStyle = string.IsNullOrEmpty(existingStyle)
                             ? newStyle
                             : existingStyle + " " + newStyle;
                    bodyNode.SetAttributeValue("style", combinedStyle);
                }



                string modifiedHtml = doc.DocumentNode.OuterHtml;
                return modifiedHtml;
            }
            else
            {
                // If it is not wrapped, wrap it with <html> and <body> tags and apply inline styles
                return $@"
                    <html>
                    <body style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>
                        {ApplyInlineStyles(originalHtml, font, fontSize)}
                    </body>
                    </html>";
            }
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
                .Replace("<li>", $"<li style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>")
                .Replace("<td>", $"<td style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>");
        }


        private string RemoveWordHeading(string html)
        {
            var result = Regex.Replace(html, @"(?s)<h\d.+?>", string.Empty);
            return Regex.Replace(result, @"</h\d>", string.Empty);
        }

        private string FixBullets(string html)
        {
            html = FixBullets(html, "MsoListParagraphCxSpFirst");
            html = FixBullets(html, "MsoListParagraphCxSpMiddle");
            html = FixBullets(html, "MsoListParagraphCxSpLast");

            return html;
        }

        private static string FixBullets(string description, string mainClassPattern)
        {
            var res = description;

            foreach (var match in Regex.Matches(description, $"(?s)<p class={mainClassPattern}.*?</p>", RegexOptions.IgnoreCase))
            {
                var bulletPattern = "(?s)<span style=\"font-family:Symbol;\">.*?</span></span></span>";

                var bulletMatch = Regex.Match(match.ToString(), bulletPattern, RegexOptions.IgnoreCase);

                if (!bulletMatch.Success)
                    continue;

                var matchWithoutBullet = Regex.Replace(match.ToString(), bulletPattern, string.Empty);

                var innerMatch = Regex.Match(matchWithoutBullet, "(?=>)(.*?)(?=</p>)", RegexOptions.Singleline);

                if (!innerMatch.Success)
                    continue;

                var newText = matchWithoutBullet.Replace(innerMatch.Value, $"><ul><li>{innerMatch.Value.Remove(0, 1)}</li></ul>");

                res = res.Replace(match.ToString(), newText);
            }

            return res;
        }

        #region HTML Validation

        /// <summary>
        /// Validates that no inline element contains a block element.
        /// Returns a list of error messages; if empty, the structure passes this rule.
        /// </summary>
        private List<string> ValidateHtmlStructure(string html)
        {
            var errors = new List<string>();

            var doc = new HtmlDocument
            {
                // Disable auto-correction so that our raw structure is examined.
                OptionFixNestedTags = false,
                OptionAutoCloseOnEnd = false
            };
            doc.LoadHtml(html);

            // Check the entire document recursively.
            ValidateNode(doc.DocumentNode, errors);

            return errors;
        }

        private void ValidateNode(HtmlNode node, List<string> errors)
        {
            // If this node is an inline element, ensure it does not contain any block element as a descendant.

            if (node.NodeType == HtmlNodeType.Element && inlineElements.Contains(node.Name))
            {

                // Descendants() gives all child nodes (recursively).
                foreach (var descendant in node.Descendants().Where(n => n.NodeType == HtmlNodeType.Element))
                {
                    if (blockElements.Contains(descendant.Name))
                    {
                        errors.Add($"Invalid nesting: Inline element <{node.Name}> contains block element <{descendant.Name}> (line {descendant.Line}, pos {descendant.LinePosition}).");
                    }
                }
            }

            if (node.NodeType == HtmlNodeType.Element && blockElements.Contains(node.Name))
            {
                if (node.Name == "ul" || node.Name == "ol")
                {
                    foreach (var childs in node.ChildNodes.Where(n => n.NodeType == HtmlNodeType.Element))
                    {
                        if (childs.Name != "li")
                        {
                            errors.Add($"Invalid nesting: Parent item <{node.Name}> cannot have <{childs.Name}> (line {childs.Line}, pos {childs.LinePosition}).");
                        }
                    }
                }
            }

            // Continue recursively.
            foreach (var child in node.ChildNodes)
            {
                ValidateNode(child, errors);
            }
        }

        #endregion

        /// <summary>
        /// Resizes images in a paragraph to fit properly in the document
        /// </summary>
        /// <param name="paragraph">The paragraph containing images to resize</param>
        private void ResizeImagesInParagraph(Paragraph paragraph)
        {
            // Find all Drawing elements (modern image format) and use PictureService to resize them
            var drawings = paragraph.Descendants<Drawing>().ToList();
            
            foreach (var drawing in drawings)
            {
                _pictureService.ResizeDrawing(drawing);
            }
            
            // Note: VML ImageData elements are legacy format and not commonly used in modern scenarios
            // If needed, VML support can be added to PictureService in the future
        }




    }
}