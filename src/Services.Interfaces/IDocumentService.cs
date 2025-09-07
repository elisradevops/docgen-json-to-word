using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace JsonToWord.Services.Interfaces
{
    public interface IDocumentService
    {
        string CreateDocument(string templatePath);
        void SetLandscape(MainDocumentPart mainPart);
        
    }
}