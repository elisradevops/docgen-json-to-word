using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    public class SectionPlaceholderService : ISectionPlaceholderService
    {
        private static readonly Regex PlaceholderRegex =
            new Regex(@"\{\{section:([0-9.]+)\}\}", RegexOptions.Compiled);

        private readonly ILogger<SectionPlaceholderService> _logger;

        public SectionPlaceholderService(ILogger<SectionPlaceholderService> logger)
        {
            _logger = logger;
        }

        /// <inheritdoc />
        public void ResolveSectionPlaceholders(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart?.Document?.Body;
            if (body == null) return;

            // Heading counters: index 0 → Heading1, index 1 → Heading2, etc.
            var counters = new int[9];
            // Track the last heading number string that was computed (e.g., "4")
            string lastHeadingNumber = "";

            foreach (var element in body.ChildElements.ToList())
            {
                if (element is Paragraph paragraph)
                {
                    var headingLevel = GetHeadingLevel(paragraph, document.MainDocumentPart);
                    if (headingLevel > 0 && headingLevel <= 9)
                    {
                        // Increment counter at this level
                        counters[headingLevel - 1]++;
                        // Reset all deeper levels
                        for (int i = headingLevel; i < 9; i++)
                            counters[i] = 0;

                        // Build the heading number string (e.g., "4" for Heading1 count=4,
                        // "4.2" for Heading2 under the 4th Heading1)
                        lastHeadingNumber = BuildHeadingNumber(counters, headingLevel);
                    }
                }
                else if (element is Table table)
                {
                    ResolveTablePlaceholders(table, lastHeadingNumber);
                }
                else if (element is SdtBlock sdtBlock)
                {
                    // Content controls may contain paragraphs (headings) or tables
                    foreach (var child in sdtBlock.Descendants<Paragraph>().ToList())
                    {
                        var headingLevel = GetHeadingLevel(child, document.MainDocumentPart);
                        if (headingLevel > 0 && headingLevel <= 9)
                        {
                            counters[headingLevel - 1]++;
                            for (int i = headingLevel; i < 9; i++)
                                counters[i] = 0;
                            lastHeadingNumber = BuildHeadingNumber(counters, headingLevel);
                        }
                    }
                    foreach (var tbl in sdtBlock.Descendants<Table>().ToList())
                    {
                        ResolveTablePlaceholders(tbl, lastHeadingNumber);
                    }
                }
            }
        }

        /// <summary>
        /// Determines the heading level (1-9) of a paragraph, or 0 if not a heading.
        /// </summary>
        private int GetHeadingLevel(Paragraph paragraph, MainDocumentPart mainPart)
        {
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrEmpty(styleId)) return 0;

            // Direct match: "Heading1" → 1, "Heading2" → 2, etc.
            if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(styleId.Substring(7), out int level)
                && level >= 1 && level <= 9)
            {
                return level;
            }

            // Check if the style is based on a heading style (e.g., custom styles)
            if (mainPart?.StyleDefinitionsPart?.Styles != null)
            {
                var style = mainPart.StyleDefinitionsPart.Styles
                    .Elements<Style>()
                    .FirstOrDefault(s => s.StyleId == styleId);

                if (style?.BasedOn?.Val != null)
                {
                    var basedOn = style.BasedOn.Val.Value;
                    if (basedOn.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                        && int.TryParse(basedOn.Substring(7), out int baseLevel)
                        && baseLevel >= 1 && baseLevel <= 9)
                    {
                        return baseLevel;
                    }
                }
            }

            return 0;
        }

        /// <summary>
        /// Builds a dotted heading number from the counters up to the given level.
        /// e.g., counters=[4,2,0,...] level=2 → "4.2"
        /// </summary>
        private string BuildHeadingNumber(int[] counters, int level)
        {
            var parts = new string[level];
            for (int i = 0; i < level; i++)
                parts[i] = counters[i].ToString();
            return string.Join(".", parts);
        }

        /// <summary>
        /// Scans all text runs in the table for {{section:X.Y}} placeholders
        /// and replaces them with parentHeading.X.Y.
        /// </summary>
        private void ResolveTablePlaceholders(Table table, string parentHeading)
        {
            if (string.IsNullOrEmpty(parentHeading)) return;

            foreach (var run in table.Descendants<Run>().ToList())
            {
                var textElement = run.GetFirstChild<Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text)) continue;

                var match = PlaceholderRegex.Match(textElement.Text);
                if (match.Success)
                {
                    var relativePath = match.Groups[1].Value;
                    var resolved = $"{parentHeading}.{relativePath}";
                    textElement.Text = PlaceholderRegex.Replace(textElement.Text, resolved);
                    _logger.LogDebug($"Resolved section placeholder: {{{{section:{relativePath}}}}} → {resolved}");
                }
            }
        }
    }
}
