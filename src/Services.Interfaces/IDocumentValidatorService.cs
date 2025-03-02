using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace JsonToWord.Services.Interfaces
{
    public interface IDocumentValidatorService
    {
        bool ValidateDocument(WordprocessingDocument document);
        List<string> ValidateInnerElementOfContentControl(string contentControlTitle, OpenXmlElement element);
    }
}
