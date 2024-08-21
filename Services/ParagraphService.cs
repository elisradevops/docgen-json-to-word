using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;

namespace JsonToWord.Services
{
    internal class ParagraphService
    {
        internal Paragraph CreateParagraph(WordParagraph wordParagraph)
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

        internal Paragraph CreateCaption(string captionText)
        {
            var run = new Run();
            run.AppendChild(new Text(captionText));

            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties();

            // Set the style of the paragraph to be a caption style (you might want to define a custom style in the Word document)
            paragraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "Caption" };

            paragraph.AppendChild(paragraphProperties);
            paragraph.AppendChild(run);

            return paragraph;
        }
    }
}