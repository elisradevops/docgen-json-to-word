using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models.TestReporterModels;

namespace JsonToWord.Services.Interfaces
{
    public interface IInternalValidationReporterService
    {
        void Insert(
            SpreadsheetDocument document,
            string worksheetName,
            InternalValidationReporterModel coverageModel
        );
    }
}
