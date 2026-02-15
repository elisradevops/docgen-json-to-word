using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models.TestReporterModels;

namespace JsonToWord.Services.Interfaces
{
    public interface IMewpCoverageReporterService
    {
        void Insert(
            SpreadsheetDocument document,
            string worksheetName,
            MewpCoverageReporterModel coverageModel
        );
    }
}
