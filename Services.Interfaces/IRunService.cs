using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using DocumentFormat.OpenXml.Wordprocessing;


namespace JsonToWord.Services.Interfaces
{
    public interface IRunService
    {
        Run CreateRun(WordRun wordRun, WordprocessingDocument document);
    }
}
