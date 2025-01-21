using JsonToWord.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace JsonToWord.Services.Interfaces
{
    public interface IParagraphService
    {
        Paragraph CreateParagraph(WordParagraph wordParagraph);
        Paragraph CreateCaption(string captionText);
    }
}
