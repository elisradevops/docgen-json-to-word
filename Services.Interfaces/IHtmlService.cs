using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using System.Threading.Tasks;

namespace JsonToWord.Services.Interfaces
{
    public interface IHtmlService
    {
        void Insert(WordprocessingDocument document, string contentControlTitle, WordHtml wordHtml);
        string CreateHtmlWordDocument(string html);
    }
}
