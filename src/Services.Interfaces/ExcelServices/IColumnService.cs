using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Models.TestReporterModels;
using System.Collections.Generic;

namespace JsonToWord.Services.Interfaces.ExcelServices
{
    public interface IColumnService
    {
        List<ColumnDefinition> DefineColumns(TestReporterModel testReporterModel, bool groupBySuite);
        Dictionary<string, int> GetColumnCountForeachGroup(List<ColumnDefinition> columnDefinitions);
        Columns CreateColumns(List<ColumnDefinition> columnDefinitions);
    }
}
