using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;

namespace JsonToWord.Services
{
    public class ParagraphService: IParagraphService
    {
        public Paragraph CreateParagraph(WordParagraph wordParagraph, bool isUnderStandardHeading)
        {
            var paragraph = new Paragraph();

            if (wordParagraph.HeadingLevel == 0)
                return paragraph;

            var headingLevel = isUnderStandardHeading ? wordParagraph.HeadingLevel : wordParagraph.HeadingLevel-1;
            var paragraphProperties = new ParagraphProperties();
            var paragraphStyleId = new ParagraphStyleId { Val = $"Heading{headingLevel}" };

            paragraphProperties.ParagraphStyleId = paragraphStyleId;
            
            // For headings under custom headings (like "Appendix"), prevent page breaks
            // This is especially important for Heading1 which often has a page break before it by default
            if (!isUnderStandardHeading)
            {
                // Add PageBreakBefore property and set it to false to prevent automatic page breaks
                paragraphProperties.AppendChild(new PageBreakBefore() { Val = false });
            }
            
            paragraph.ParagraphProperties = paragraphProperties;

            return paragraph;
        }


        public Paragraph CreateCaption(string captionText)
        {
            Paragraph paragraph1 = new Paragraph();
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Caption" };
            Justification justification1 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties1.ParagraphStyleId = paragraphStyleId1;
            paragraphProperties1.Justification = justification1;
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = captionText;

            run1.Append(text1);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };
            paragraph1.ParagraphProperties = paragraphProperties1;
            paragraph1.Append(proofError1);
            paragraph1.Append(run1);
            paragraph1.Append(proofError2);
            return paragraph1;
        }

        /// <summary>
        /// Applies tight spacing to a paragraph by reducing line spacing and paragraph spacing
        /// </summary>
        /// <param name="paragraph">The paragraph to apply tight spacing to</param>
        public void ApplyTightSpacing(Paragraph paragraph)
        {
            // Get or create paragraph properties
            if (paragraph.ParagraphProperties == null)
            {
                paragraph.ParagraphProperties = new ParagraphProperties();
            }

            // Check if paragraph contains images or drawings
            bool hasImages = paragraph.Descendants<Drawing>().Any() || 
                           paragraph.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any();

            SpacingBetweenLines spacingBetweenLines;
            
            if (hasImages)
            {
                // For paragraphs with images, use automatic spacing but reduce paragraph spacing
                spacingBetweenLines = new SpacingBetweenLines()
                {
                    LineRule = LineSpacingRuleValues.Auto,  // Auto spacing for images
                    Line = "240",        // Single line spacing (1.0)
                    Before = "0",        // No space before paragraph
                    After = "60"         // Minimal space after paragraph (3pt)
                };
                

            }
            else
            {
                // For text-only paragraphs, use tighter exact spacing
                spacingBetweenLines = new SpacingBetweenLines()
                {
                    Line = "240",        // 240 twentieths of a point (12pt line height)
                    LineRule = LineSpacingRuleValues.Exact,
                    Before = "0",        // No space before paragraph
                    After = "0"          // No space after paragraph
                };
            }

            paragraph.ParagraphProperties.SpacingBetweenLines = spacingBetweenLines;

            // Set paragraph margins to zero for tighter spacing
            var indentation = new Indentation()
            {
                Left = "0",
                Right = "0",
                FirstLine = "0"
            };

            paragraph.ParagraphProperties.Indentation = indentation;
        }

    }
}