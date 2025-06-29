using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace JsonToWord.Services.Interfaces.ExcelServices
{
    // IStylesheetService.cs - Handles styles, fonts, colors
    public interface IStylesheetService
    {
        void EnsureStylesheet(WorkbookPart workbookPart);
        Stylesheet CreateStylesheet();
        // Other style-related methods
    }
}
