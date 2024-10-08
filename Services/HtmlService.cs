﻿using System;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using JsonToWord.Models;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    internal class HtmlService
    {
        private readonly ContentControlService _contentControlService;
        private readonly ILogger<HtmlService> _logger;
        public HtmlService()
        {
            _contentControlService = new ContentControlService();
        }
        internal void Insert(WordprocessingDocument document, string contentControlTitle, WordHtml wordHtml)
        {
            var html = SetHtmlFormat(wordHtml.Html, wordHtml.Font, wordHtml.FontSize);
            
            html = RemoveWordHeading(html);

            html = FixBullets(html);

            var tempHtmlFile = CreateHtmlWordDocument(html);

            var altChunkId = "altChunkId" + Guid.NewGuid().ToString("N");
            var mainPart = document.MainDocumentPart;
            var chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            using (var fileStream = File.Open(tempHtmlFile, FileMode.Open))
            {
                chunk.FeedData(fileStream);
            }

            var altChunk = new AltChunk { Id = altChunkId };
            
            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(altChunk);

            sdtBlock.AppendChild(sdtContentBlock);
        }

        internal string CreateHtmlWordDocument(string html)
        {
            var tempHtmlDirectory = Path.Combine(Path.GetTempPath(), "MicrosoftWordOpenXml", Guid.NewGuid().ToString("N"));

            if (!Directory.Exists(tempHtmlDirectory))
                Directory.CreateDirectory(tempHtmlDirectory);

            var tempHtmlFile = CreateTempDocument(tempHtmlDirectory);

            using (var document = WordprocessingDocument.Open(tempHtmlFile, true))
            {
                var mainPart = document.MainDocumentPart;

                if (mainPart == null)
                {
                    mainPart = document.AddMainDocumentPart();
                    new Document(new Body()).Save(mainPart);
                }

                var converter = new HtmlConverter(mainPart);
                try
                {
                    converter.ParseHtml(html);
                }
                catch(Exception ex)
                {
                    _logger.LogError("DocGen ran into an issue parsing the html due to :" , ex);
                    converter.ParseHtml("<p style='color: red'><b>DocGen ran into an issue parsing the html due to :" + ex.Message +"<b></p>");
                }
            }

            return tempHtmlFile;
        }

        private string CreateTempDocument(string directory)
        {
            var tempDocumentFile = Path.Combine(directory, $"{Guid.NewGuid():N}.docx");

            using (var wordDocument = WordprocessingDocument.Create(tempDocumentFile, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();

                mainPart.Document = new Document();
                mainPart.Document.AppendChild(new Body());
            }

            return tempDocumentFile;
        }

        private string SetHtmlFormat(string html, string font, uint fontSize)
        {
            if (!html.ToLower().StartsWith("<html>"))
            {
                // This method wraps the HTML content with inline styles, since Word does not reliably support <style> tags in altChunk
                return $@"
                    <html>
                    <body style='font-family: {font}, sans-serif; font-size: {fontSize}pt;'>
                        {ApplyInlineStyles(html, font, fontSize)}
                    </body>
                    </html>";
            }

            return html;
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