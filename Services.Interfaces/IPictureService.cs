using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;

namespace JsonToWord.Services.Interfaces
{
    public interface IPictureService
    {
        void Insert(WordprocessingDocument document, string contentControlTitle, WordAttachment wordAttachment);
        Drawing CreateDrawing(MainDocumentPart mainDocumentPart, string filePath, bool isFlattened = false);
    }
}
