using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using HtmlToOpenXml;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    internal class HtmlService : IHtmlService
    {
        private readonly IContentControlService _contentControlService;
        private readonly IDocumentValidatorService _documentValidator;
        private readonly ILogger<HtmlService> _logger;
        public HtmlService(IContentControlService contentControlService,IDocumentValidatorService documentValidator, ILogger<HtmlService> logger)
        {
            _contentControlService = contentControlService;
            _logger = logger;
            _documentValidator = documentValidator;
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

            try
            {
                var elements = converter.Parse(html);
                return elements;
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
                _logger.LogError(ex, $"DocGen ran into an issue parsing the html due to: {errorMessage}");
                _logger.LogError($"The html that caused the issue is: {html}");

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
    }
}