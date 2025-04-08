using JsonToWord.Models;

namespace JsonToWord.Services.Interfaces
{
    public interface IExcelService
    {
        string CreateExcelDocument(ExcelModel excelModel);
    }
}
