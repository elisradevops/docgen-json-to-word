using DocumentFormat.OpenXml.Packaging;

namespace JsonToWord.Services.Interfaces
{
    public interface IUtilsService
    {
        int ConvertDxaToPct(int dxa, int pageWidthDxa);
        int ConvertCmToDxa(double cm);
        int GetPageWidthDxa(MainDocumentPart mainPart);
        double ParseStringToDouble(string input);
    }
}
