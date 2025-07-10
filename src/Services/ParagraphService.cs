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
    }
}