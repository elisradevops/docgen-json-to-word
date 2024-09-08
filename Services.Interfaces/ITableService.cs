using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;

namespace JsonToWord.Services.Interfaces
{
    public interface ITableService
    {
        void Insert(WordprocessingDocument document, string contentControlTitle, WordTable wordTable);

    }
}
