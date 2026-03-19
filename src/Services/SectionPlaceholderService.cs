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
        private const string SysRsAnchorName = "requirements-root";
        private static readonly Regex PlaceholderRegex =
            new Regex(@"\{\{section:(?:(?<anchor>[A-Za-z0-9_-]+):)?(?<path>[0-9.]+)\}\}", RegexOptions.Compiled);
        private static readonly Regex AnchorMarkerRegex =
            new Regex(@"\{\{section-anchor:(?<anchor>[A-Za-z0-9_-]+)\}\}", RegexOptions.Compiled);

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
            var anchorHeadingMap = new System.Collections.Generic.Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase
            );

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

                    CaptureAndClearAnchorMarkers(paragraph, lastHeadingNumber, anchorHeadingMap);
                }
                else if (element is Table table)
                {
                    CaptureAndClearAnchorMarkers(table, lastHeadingNumber, anchorHeadingMap);
                    ResolveTablePlaceholders(table, lastHeadingNumber, anchorHeadingMap);
                }
                else if (element is SdtBlock sdtBlock)
                {
                    CaptureAndClearAnchorMarkers(sdtBlock, lastHeadingNumber, anchorHeadingMap);

                    // Content controls contain generated headings (e.g., requirement
                    // hierarchy) that must NOT affect the document-level heading counters.
                    // Only resolve table placeholders using the last heading number
                    // computed from the static template headings above.
                    foreach (var tbl in sdtBlock.Descendants<Table>().ToList())
                    {
                        ResolveTablePlaceholders(tbl, lastHeadingNumber, anchorHeadingMap);
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
        private void ResolveTablePlaceholders(
            Table table,
            string parentHeading,
            System.Collections.Generic.IDictionary<string, string> anchorHeadingMap
        )
        {
            foreach (var run in table.Descendants<Run>().ToList())
            {
                var textElement = run.GetFirstChild<Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text)) continue;

                var originalText = textElement.Text;
                var replacedText = PlaceholderRegex.Replace(originalText, (match) =>
                {
                    var anchor = match.Groups["anchor"]?.Value;
                    var relativePath = match.Groups["path"]?.Value;

                    var baseHeading = parentHeading;
                    var hasAnchoredBase = false;
                    if (!string.IsNullOrEmpty(anchor)
                        && anchorHeadingMap != null
                        && anchorHeadingMap.TryGetValue(anchor, out var anchoredHeading)
                        && !string.IsNullOrEmpty(anchoredHeading))
                    {
                        baseHeading = anchoredHeading;
                        hasAnchoredBase = true;
                    }

                    if (string.IsNullOrEmpty(baseHeading) || string.IsNullOrEmpty(relativePath))
                    {
                        return match.Value;
                    }

                    var resolved = hasAnchoredBase
                        ? ResolveAnchoredSection(anchor, baseHeading, relativePath)
                        : $"{baseHeading}.{relativePath}";

                    if (string.IsNullOrEmpty(resolved))
                    {
                        return match.Value;
                    }

                    _logger.LogDebug(
                        $"Resolved section placeholder: {match.Value} → {resolved} (anchor: {(string.IsNullOrEmpty(anchor) ? "<parent>" : anchor)})"
                    );
                    return resolved;
                });

                if (!string.Equals(originalText, replacedText, StringComparison.Ordinal))
                {
                    textElement.Text = replacedText;
                }
            }
        }

        private string ResolveAnchoredSection(string anchorName, string anchorHeading, string relativePath)
        {
            var anchorParts = anchorHeading.Split('.');
            var pathParts = relativePath.Split('.');
            if (anchorParts.Length == 0 || pathParts.Length == 0)
            {
                return string.Empty;
            }

            if (!int.TryParse(pathParts[0], out var firstRelativeNumber))
            {
                return string.Empty;
            }

            if (string.Equals(anchorName, SysRsAnchorName, StringComparison.OrdinalIgnoreCase))
            {
                // SysRS scanning model: section numbering is chapter-based.
                // We anchor on Chapter(H1) and advance the next H2 index, ignoring deeper
                // headings before the marker (e.g., 4.1.2 + 1 => 4.2).
                if (!int.TryParse(anchorParts[0], out var chapterNumber))
                {
                    return string.Empty;
                }

                var existingSectionOffset = 0;
                if (anchorParts.Length > 1 && int.TryParse(anchorParts[1], out var existingSecondLevel))
                {
                    existingSectionOffset = existingSecondLevel;
                }

                var resolvedParts = new string[pathParts.Length + 1];
                resolvedParts[0] = chapterNumber.ToString();
                resolvedParts[1] = (existingSectionOffset + firstRelativeNumber).ToString();

                for (int i = 1; i < pathParts.Length; i++)
                {
                    if (!int.TryParse(pathParts[i], out var parsedPart))
                    {
                        return string.Empty;
                    }

                    resolvedParts[i + 1] = parsedPart.ToString();
                }

                return string.Join(".", resolvedParts);
            }

            // Generic anchored behavior for non-SysRS anchors:
            // continue numbering from the anchor depth.
            if (anchorParts.Length == 1)
            {
                return $"{anchorHeading}.{relativePath}";
            }

            var genericParts = new string[(anchorParts.Length - 1) + pathParts.Length];
            for (int i = 0; i < anchorParts.Length - 1; i++)
            {
                if (!int.TryParse(anchorParts[i], out var anchorPart))
                {
                    return string.Empty;
                }

                genericParts[i] = anchorPart.ToString();
            }

            if (!int.TryParse(anchorParts[anchorParts.Length - 1], out var anchorLastPart))
            {
                return string.Empty;
            }

            genericParts[anchorParts.Length - 1] = (anchorLastPart + firstRelativeNumber).ToString();

            for (int i = 1; i < pathParts.Length; i++)
            {
                if (!int.TryParse(pathParts[i], out var parsedPart))
                {
                    return string.Empty;
                }

                genericParts[(anchorParts.Length - 1) + i] = parsedPart.ToString();
            }

            return string.Join(".", genericParts);
        }

        private void CaptureAndClearAnchorMarkers(
            OpenXmlElement container,
            string currentHeading,
            System.Collections.Generic.IDictionary<string, string> anchorHeadingMap
        )
        {
            if (container == null || string.IsNullOrEmpty(currentHeading) || anchorHeadingMap == null)
            {
                return;
            }

            foreach (var textElement in container.Descendants<Text>().ToList())
            {
                if (string.IsNullOrEmpty(textElement.Text)) continue;

                var originalText = textElement.Text;
                foreach (Match match in AnchorMarkerRegex.Matches(originalText))
                {
                    var anchor = match.Groups["anchor"]?.Value;
                    if (string.IsNullOrWhiteSpace(anchor)) continue;
                    anchorHeadingMap[anchor] = currentHeading;
                    _logger.LogDebug($"Captured section anchor marker: {anchor} → {currentHeading}");
                }

                var cleanedText = AnchorMarkerRegex.Replace(originalText, string.Empty);
                if (!string.Equals(originalText, cleanedText, StringComparison.Ordinal))
                {
                    textElement.Text = cleanedText;
                }
            }
        }
    }
}
