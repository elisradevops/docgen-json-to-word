using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;

namespace JsonToWord.Services.Interfaces
{
    public interface ITextService
    {
        void Write(WordprocessingDocument document, string contentControlTitle, WordParagraph wordParagraph);

    }
}
