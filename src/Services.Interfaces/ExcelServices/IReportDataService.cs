using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Models.TestReporterModels;
using System.Collections.Generic;

namespace JsonToWord.Services.Interfaces.ExcelServices
{
    public interface IReportDataService
    {
        void AddDataRows(SheetData sheetData, MergeCells mergeCells,
                List<TestSuiteModel> testSuites,
                List<ColumnDefinition> columnDefinitions,
                Dictionary<string, int> columnCountForeachGroup,
                ref uint rowIndex, WorksheetPart worksheetPart, bool groupBySuite);
    }
}
