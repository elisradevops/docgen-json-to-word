using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using System;

namespace JsonToWord.Services
{
    public class ParagraphService: IParagraphService
    {
        public Paragraph CreateParagraph(WordParagraph wordParagraph)
        {
            var paragraph = new Paragraph();

            if (wordParagraph.HeadingLevel == 0)
                return paragraph;

            var paragraphProperties = new ParagraphProperties();
            var paragraphStyleId = new ParagraphStyleId { Val = $"Heading{wordParagraph.HeadingLevel}" };

            paragraphProperties.AppendChild(paragraphStyleId);
            paragraph.AppendChild(paragraphProperties);

            return paragraph;
        }



        public Paragraph InitParagraphForListItem(
            WordListItem wordListItem,
            bool isOrdered,
            int numberingId,
            bool multiLevel)
        {
            var paragraph = new Paragraph();
            var pPr = new ParagraphProperties();

            // If single-level => clamp to 0
            int level = multiLevel ? wordListItem.Level : 0;
            level = Math.Max(0, Math.Min(level, 8)); // clamp

            var numProps = new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = numberingId }
            );
            pPr.Append(numProps);

            pPr.ParagraphStyleId = new ParagraphStyleId { Val = "ListParagraph" };

            paragraph.ParagraphProperties = pPr;
            return paragraph;
        }



        public Paragraph CreateCaption(string captionText)
        {
            Paragraph paragraph1 = new Paragraph();
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Caption" };
            Justification justification1 = new Justification() { Val = JustificationValues.Left };


            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(justification1);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = captionText;

            run1.Append(text1);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(proofError1);
            paragraph1.Append(run1);
            paragraph1.Append(proofError2);
            return paragraph1;
        }
    }
}