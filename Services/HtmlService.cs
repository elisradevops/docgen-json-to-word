using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using HtmlToOpenXml;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using SixLabors.ImageSharp;

namespace JsonToWord.Services
{
    internal class HtmlService : IHtmlService
    {
        private readonly ContentControlService _contentControlService;
        private readonly ILogger<HtmlService> _logger;
        public HtmlService(ILogger<HtmlService> logger)
        {
            _contentControlService = new ContentControlService();
            _logger = logger;
        }
        public void Insert(WordprocessingDocument document, string contentControlTitle, WordHtml wordHtml)
        {
            var elements = ConvertHtmlToOpenXmlElements(wordHtml, document);

            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();

            sdtContentBlock.Append(elements);

            sdtBlock.AppendChild(sdtContentBlock);
        }


        public IEnumerable<OpenXmlCompositeElement> ConvertHtmlToOpenXmlElements(WordHtml wordHtml, WordprocessingDocument document)
        {
            var html = WrapHtmlWithStyle(wordHtml.Html, wordHtml.Font, wordHtml.FontSize);

            html = RemoveWordHeading(html);

            html = FixBullets(html);
            var converter = new HtmlConverter(document.MainDocumentPart, new HtmlToOpenXml.IO.DefaultWebRequest()
            {
                BaseImageUrl = new Uri(Environment.CurrentDirectory)
            });
            converter.ContinueNumbering = false;
            converter.SupportsHeadingNumbering = false;

            var elements = converter.Parse(html);

            return elements;
        }

        private void FixNumberingIdConflicts(NumberingDefinitionsPart numberingPart)
        {
            var numbering = numberingPart.Numbering;
            if (numbering == null)
                return;

            // Collect existing abstractNumIds and numIds
            var abstractNumIds = numbering.Elements<AbstractNum>()
                .Select(an => an.AbstractNumberId.Value)
                .ToList();

            var numIds = numbering.Elements<NumberingInstance>()
                .Select(ni => ni.NumberID.Value)
                .ToList();

            // Create mappings to keep track of new IDs
            var abstractNumIdMapping = new Dictionary<int, int>();
            var numIdMapping = new Dictionary<int, int>();

            int nextAbstractNumId = abstractNumIds.Any() ? abstractNumIds.Max() + 1 : 1;
            int nextNumId = numIds.Any() ? numIds.Max() + 1 : 1;

            // Ensure unique abstractNumIds and map levels
            foreach (var abstractNum in numbering.Elements<AbstractNum>())
            {
                int oldId = abstractNum.AbstractNumberId.Value;
                if (abstractNumIds.Count(id => id == oldId) > 1 || abstractNumIdMapping.ContainsKey(oldId))
                {
                    abstractNum.AbstractNumberId.Value = nextAbstractNumId;
                    abstractNumIdMapping[oldId] = nextAbstractNumId;
                    nextAbstractNumId++;
                }
                else
                {
                    abstractNumIdMapping[oldId] = oldId;
                }

                // Assign numbering formats for different levels
                foreach (var level in abstractNum.Elements<Level>())
                {
                    switch (level.LevelIndex.Value)
                    {
                        case 0:
                            level.NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                            break;
                        case 1:
                            level.NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                            break;
                        case 2:
                            level.NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                            break;
                        default:
                            level.NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                            break;
                    }
                }
            }

            // Ensure unique numIds and update AbstractNumId references
            foreach (var numberingInstance in numbering.Elements<NumberingInstance>())
            {
                int oldNumId = numberingInstance.NumberID.Value;
                if (numIds.Count(id => id == oldNumId) > 1 || numIdMapping.ContainsKey(oldNumId))
                {
                    numberingInstance.NumberID.Value = nextNumId;
                    numIdMapping[oldNumId] = nextNumId;
                    nextNumId++;
                }
                else
                {
                    numIdMapping[oldNumId] = oldNumId;
                }

                // Update the AbstractNumId reference
                int oldAbstractNumId = numberingInstance.AbstractNumId.Val.Value;
                numberingInstance.AbstractNumId.Val.Value = abstractNumIdMapping[oldAbstractNumId];
            }

            // Save changes to the numbering definitions part
            numbering.Save(numberingPart);
        }
        private void AssertThatHtmlToOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
        {
            var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
            var errors = validator.Validate(wpDoc);

            if (!errors.GetEnumerator().MoveNext())
                return;

            var errorMessage = new StringBuilder("The document doesn't look 100% compatible with Office 2016.\n");

            foreach (ValidationErrorInfo error in errors)
            {
                errorMessage.AppendFormat("{0}\n\t{1}\n", error.Path.XPath, error.Description);
            }

            throw new InvalidOperationException(errorMessage.ToString());
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
    }
}