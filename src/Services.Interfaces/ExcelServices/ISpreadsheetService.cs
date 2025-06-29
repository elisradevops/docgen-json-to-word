using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System;
using JsonToWord.Models.Excel;

namespace JsonToWord.Services.Interfaces.ExcelServices
{
    public interface ISpreadsheetService
    {
        WorksheetPart GetOrCreateWorksheetPart(WorkbookPart workbookPart, string worksheetName);
        void CreateHeaderRow(SheetData sheetData, List<ColumnDefinition> columnDefinitions, MergeCells mergeCells, Dictionary<string, int> columnCountForeachGroup, Dictionary<string, int> groupItemCounts);
        Cell CreateTextCell(string cellReference, string cellValue, uint styleIndex = 0);
        Cell CreateNumberCell(string cellReference, string cellValue, uint styleIndex = 0);
        Cell CreateDateCell(string cellReference, DateTime date, uint styleIndex = 0);
        Cell CreateHyperlinkCell(WorksheetPart worksheetPart, string cellReference, string displayText, string url, uint styleIndex, string tooltipMessage);
        string GetColumnLetter(int columnIndex);
    }
}
