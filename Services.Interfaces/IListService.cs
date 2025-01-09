using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;

namespace JsonToWord.Services.Interfaces
{
    public interface IListService
    {
        void Insert(WordprocessingDocument document, string contentControlTitle, WordList wordList);
    }
}
