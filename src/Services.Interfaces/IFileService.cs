using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.EventHandlers;
using JsonToWord.Models;

namespace JsonToWord.Services.Interfaces
{
    public interface IFileService
    {
        void Insert(WordprocessingDocument document, string contentControlTitle, WordAttachment wordAttachment);
        Paragraph AttachFileToParagraph(MainDocumentPart mainPart, WordAttachment wordAttachment);
        event NonOfficeAttachmentEventHandler nonOfficeAttachmentEventHandler;

    }
}
