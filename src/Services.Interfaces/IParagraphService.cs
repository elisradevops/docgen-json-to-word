using JsonToWord.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace JsonToWord.Services.Interfaces
{
    public interface IParagraphService
    {
        Paragraph CreateParagraph(WordParagraph wordParagraph, bool isUnderStandardHeading);
        Paragraph CreateCaption(string captionText);
    }
}
