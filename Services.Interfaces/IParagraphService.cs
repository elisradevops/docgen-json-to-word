using JsonToWord.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace JsonToWord.Services.Interfaces
{
    public interface IParagraphService
    {
        Paragraph CreateParagraph(WordParagraph wordParagraph);
        Paragraph InitParagraphForListItem(WordListItem wordListItem, bool isOrdered, int numberingId, bool multiLevel);
        Paragraph CreateCaption(string captionText);
    }
}
