using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonToWord.Models.Excel;
using JsonToWord.Services.Interfaces.ExcelServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace JsonToWord.Services.ExcelServices
{
    public class SpreadsheetService : ISpreadsheetService
    {
        public Cell CreateDateCell(string cellReference, DateTime date, uint styleIndex = 0)
        {
            return new Cell
            {
                CellReference = cellReference,
                CellValue = new CellValue(date.ToOADate().ToString(CultureInfo.InvariantCulture)),
                StyleIndex = styleIndex,
                DataType = CellValues.Number
            };
        }

        public void CreateHeaderRow(SheetData sheetData, List<ColumnDefinition> columnDefinitions, MergeCells mergeCells, Dictionary<string, int> columnCountForeachGroup, Dictionary<string, int> groupItemCounts)
        {
            uint rowIndex = 1;

            // Create the styled and merged group header row
            CreateGroupHeadersRow(sheetData, mergeCells, ref rowIndex, columnCountForeachGroup, groupItemCounts);

            // Create the field header row below it
            Row fieldHeaderRow = new Row { RowIndex = rowIndex };
            sheetData.Append(fieldHeaderRow);

            for (int i = 0; i < columnDefinitions.Count; i++)
            {
                string columnLetter = GetColumnLetter(i + 1);
                Cell fieldHeaderCell = CreateTextCell(
                    $"{columnLetter}{rowIndex}",
                    columnDefinitions[i].Name,
                    1 // General Header Style
                );
                fieldHeaderRow.Append(fieldHeaderCell);
            }
        }

        public void CreateGroupHeadersRow(SheetData sheetData, MergeCells mergeCells, ref uint rowIndex, Dictionary<string, int> columnCountForeachGroup, Dictionary<string, int> groupItemCounts)
        {
            Row row = new Row { RowIndex = rowIndex };
            sheetData.Append(row);

            uint currentColumn = 1;
            foreach (var group in columnCountForeachGroup)
            {
                string groupName = group.Key;
                int columnSpan = group.Value;

                if (columnSpan > 0)
                {
                    uint styleIndex = GetGroupHeaderStyleIndex(groupName);

                    // Get the count for the current group, defaulting to 0 if not found
                    groupItemCounts.TryGetValue(groupName, out int itemCount);
                    string headerText = itemCount > 0 ? $"{groupName} ({itemCount})" : groupName;

                    Cell cell = CreateTextCell($"{GetColumnLetter((int)currentColumn)}{row.RowIndex}", headerText, styleIndex);
                    row.Append(cell);

                    // Fill in the rest of the cells for the merge to apply style correctly
                    for (int i = 1; i < columnSpan; i++)
                    {
                        row.Append(CreateTextCell($"{GetColumnLetter((int)currentColumn + i)}{row.RowIndex}", string.Empty, styleIndex));
                    }

                    // Add the merge cell instruction
                    if (columnSpan > 1)
                    {
                        string startColLetter = GetColumnLetter((int)currentColumn);
                        string endColLetter = GetColumnLetter((int)currentColumn + columnSpan - 1);
                        mergeCells.Append(new MergeCell
                        {
                            Reference = new StringValue($"{startColLetter}{row.RowIndex}:{endColLetter}{row.RowIndex}")
                        });
                    }

                    currentColumn += (uint)columnSpan;
                }
            }
            rowIndex++; // Increment the row index for the next row
        }

        public Cell CreateHyperlinkCell(WorksheetPart worksheetPart, string cellReference, string displayText, string url, uint styleIndex, string tooltipMessage)
        {
            // Create a hyperlink relationship in the worksheet
            Uri uri = new Uri(url, UriKind.Absolute);

            // Ensure the Worksheet object exists
            if (worksheetPart.Worksheet == null)
                worksheetPart.Worksheet = new Worksheet();

            // Create the hyperlink relationship
            HyperlinkRelationship hyperlinkRelationship = worksheetPart.AddHyperlinkRelationship(uri, true);

            // Map the original style to the corresponding hyperlink style
            uint hyperlinkStyleIndex;
            if (styleIndex == 6 || styleIndex == 8 || styleIndex == 10) // First alternating color (text, date, number)
                hyperlinkStyleIndex = 12;  // Use hyperlink style for first alternating color
            else if (styleIndex == 7 || styleIndex == 9 || styleIndex == 11) // Second alternating color (text, date, number)
                hyperlinkStyleIndex = 13;  // Use hyperlink style for second alternating color
            else
                hyperlinkStyleIndex = 12;  // Default to first hyperlink style

            // Create the cell with the hyperlink text and proper style
            Cell cell = new Cell
            {
                CellReference = cellReference,
                DataType = CellValues.String, // Always use string for hyperlinks
                StyleIndex = hyperlinkStyleIndex,
                CellValue = new CellValue(displayText)
            };

            // Create the hyperlink and add it to the worksheet
            Hyperlink hyperlink = new Hyperlink
            {
                Reference = cellReference,
                Id = hyperlinkRelationship.Id,
                Tooltip = tooltipMessage
            };

            // Find or create Hyperlinks element in the worksheet
            Hyperlinks hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null)
            {
                hyperlinks = new Hyperlinks();
                worksheetPart.Worksheet.Append(hyperlinks);
            }

            hyperlinks.Append(hyperlink);

            return cell;
        }

        public Cell CreateNumberCell(string cellReference, string cellValue, uint styleIndex = 0)
        {
            return new Cell
            {
                CellReference = cellReference,
                CellValue = new CellValue(cellValue),
                DataType = CellValues.Number,
                StyleIndex = styleIndex
            };
        }

        private uint GetGroupHeaderStyleIndex(string groupName)
        {
            return groupName switch
            {
                "Test Cases" => 14,
                "Requirements" => 15,
                "Bugs" => 16,
                "CRs" => 17,
                _ => 1, // Default header style
            };
        }

        public Cell CreateTextCell(string cellReference, string cellValue, uint styleIndex = 0)
        {
            return new Cell
            {
                CellReference = cellReference,
                CellValue = new CellValue(cellValue ?? ""),
                DataType = CellValues.String,
                StyleIndex = styleIndex
            };
        }

        public string GetColumnLetter(int columnIndex)
        {
            string columnName = string.Empty;
            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                columnName = Convert.ToChar('A' + remainder) + columnName;
                columnIndex = (columnIndex - remainder - 1) / 26;
            }
            return columnName;
        }

        public WorksheetPart GetOrCreateWorksheetPart(WorkbookPart workbookPart, string worksheetName)
        {
            // Remove any invalid characters from the worksheet name
            string safeWorksheetName = GetSafeWorksheetName(worksheetName);

            // Try to find existing worksheet
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name, safeWorksheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet != null)
            {
                // Return existing worksheet part
                return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            }
            else
            {
                // Create new worksheet part
                WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add to workbook
                uint sheetId = (uint)(workbookPart.Workbook.Sheets?.Count() + 1 ?? 1);
                Sheet newSheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = sheetId,
                    Name = safeWorksheetName
                };

                if (workbookPart.Workbook.Sheets == null)
                {
                    workbookPart.Workbook.AppendChild(new Sheets());
                }

                workbookPart.Workbook.Sheets.Append(newSheet);
                return newWorksheetPart;
            }
        }

        private string GetSafeWorksheetName(string name)
        {
            // Excel worksheet name rules:
            // - Max 31 characters
            // - Cannot contain: \ / ? * [ or ]
            // - Cannot be blank
            // - Cannot start or end with an apostrophe
            if (string.IsNullOrWhiteSpace(name))
                return "Sheet1";

            // Remove invalid characters
            string invalidChars = @"\/*?[]'";
            string safeName = new string(name
                .Where(c => !invalidChars.Contains(c))
                .ToArray())
                .Trim();

            // Trim to max length
            return safeName.Length > 31 ? safeName.Substring(0, 31) : safeName;
        }
    }
}
