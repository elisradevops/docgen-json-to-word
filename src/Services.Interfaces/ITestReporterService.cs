using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using JsonToWord.Models.TestReporterModels;

namespace JsonToWord.Services.Interfaces
{
    public interface ITestReporterService
    {
        void Insert(SpreadsheetDocument document, string contentControlTitle, TestReporterModel testReporterModel);

    }
}
