using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models.TestReporterModels;

namespace JsonToWord.Services.Interfaces
{
    public interface IFlatTestReporterService
    {
        void Insert(SpreadsheetDocument document, string worksheetName, FlatTestReporterModel flatReportModel);
    }
}
