using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace JsonToWord.Services.Interfaces
{
    public interface IHtmlService
    {
        void Insert(WordprocessingDocument document, string contentControlTitle, WordHtml wordHtml, FormattingSettings formattingSettings);
        IEnumerable<OpenXmlCompositeElement> ConvertHtmlToOpenXmlElements(WordHtml wordHtml, WordprocessingDocument document);
    }
}
