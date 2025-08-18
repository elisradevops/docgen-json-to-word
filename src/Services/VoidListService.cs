using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using JsonToWord.Services.Interfaces;
using JsonToWord.Services.Interfaces.ExcelServices;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace JsonToWord.Services
{
    public class VoidListService : IVoidListService
    {
        private readonly ILogger<VoidListService> _logger;
        private readonly ISpreadsheetService _spreadsheetService;
        private static readonly Regex vlRegex = new Regex(@"#VL-[^#]+#", RegexOptions.IgnoreCase);

        public VoidListService(ILogger<VoidListService> logger, ISpreadsheetService spreadsheetService)
        {
            _logger = logger;
            _spreadsheetService = spreadsheetService;
        }

        public string CreateVoidList(string docPath)
        {
            string docName = Path.GetFileName(docPath);
            string voidListFile = Path.Combine(Path.GetDirectoryName(docPath) ?? string.Empty, docName + " - VOID LIST.xlsx")?.Replace(':', '_');
            
            var allMatches = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docPath, true))
                {
                    var mainPart = wordDoc.MainDocumentPart;
                    if (mainPart?.Document?.Body == null)
                    {
                        _logger.LogWarning("Document body is null. Cannot process for VOID list.");
                        return string.Empty;
                    }

                    foreach (var p in mainPart.Document.Body.Descendants<Paragraph>().ToList())
                    {
                        var runsToProcess = p.Elements<Word.Run>().Where(r => vlRegex.IsMatch(r.InnerText)).ToList();
                        foreach (var run in runsToProcess)
                        {
                            var newRuns = new List<Word.Run>();
                            string runText = run.InnerText;
                            var matches = vlRegex.Matches(runText);
                            int lastIndex = 0;

                            foreach (Match match in matches)
                            {
                                // Add the text before the match
                                if (match.Index > lastIndex)
                                {
                                    string beforeText = runText.Substring(lastIndex, match.Index - lastIndex);
                                    var beforeRun = new Word.Run(new Word.Text(beforeText) { Space = SpaceProcessingModeValues.Preserve });
                                    if (run.RunProperties != null) beforeRun.RunProperties = (Word.RunProperties)run.RunProperties.CloneNode(true);
                                    newRuns.Add(beforeRun);
                                }

                                // Transform and add the formatted match
                                string originalMatchValue = match.Value;
                                // In-document replacement logic
                                string[] parts = originalMatchValue.Trim('#').Split(new[] { ' ' }, 2);
                                string key = parts[0].ToUpper();
                                string newMatchValue = "#" + key;

                                // Data collection for Excel
                                string value = parts.Length > 1 ? parts[1] : string.Empty;
                                allMatches[key] = value; // Add to dictionary

                                var matchRun = new Word.Run(new Word.Text(newMatchValue) { Space = SpaceProcessingModeValues.Preserve });
                                Word.RunProperties rp = (run.RunProperties != null) ? (Word.RunProperties)run.RunProperties.CloneNode(true) : new Word.RunProperties();
                                rp.Append(new Word.Bold());
                                rp.Append(new Word.Color() { Val = "0000FF" });
                                rp.Append(new Word.Underline() { Val = Word.UnderlineValues.Single });
                                matchRun.RunProperties = rp;
                                newRuns.Add(matchRun);

                                lastIndex = match.Index + match.Length;
                            }

                            // Add any remaining text after the last match
                            if (lastIndex < runText.Length)
                            {
                                string afterText = runText.Substring(lastIndex);
                                var afterRun = new Word.Run(new Word.Text(afterText) { Space = SpaceProcessingModeValues.Preserve });
                                if (run.RunProperties != null) afterRun.RunProperties = (Word.RunProperties)run.RunProperties.CloneNode(true);
                                newRuns.Add(afterRun);
                            }

                            // Replace the old run with the new set of runs
                            foreach (var newRun in newRuns)
                            {
                                p.InsertBefore(newRun, run);
                            }
                            run.Remove();
                        }
                    }

                    mainPart.Document.Save();
                }

                if (allMatches.Count == 0)
                {
                    _logger.LogInformation("No VOID list matches found in the document.");
                    return string.Empty;
                }

                using (var spreadsheetDocument = SpreadsheetDocument.Create(voidListFile, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                                        SheetViews sheetViews = new SheetViews(
                        new SheetView() { WorkbookViewId = 0, RightToLeft = false }
                    );
                    worksheetPart.Worksheet = new Worksheet(sheetViews, new SheetData());

                    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "VOID List" };
                    sheets.Append(sheet);

                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    // Add header row
                    Row headerRow = new Row() { RowIndex = 1 };
                    headerRow.Append(
                        _spreadsheetService.CreateTextCell("A1", "VL Code"),
                        _spreadsheetService.CreateTextCell("B1", "Content")
                    );
                    sheetData.Append(headerRow);

                    // Add data rows
                    uint rowIndex = 2;
                    foreach (var entry in allMatches)
                    {
                        Row dataRow = new Row() { RowIndex = rowIndex };
                        dataRow.Append(
                            _spreadsheetService.CreateTextCell($"A{rowIndex}", entry.Key),
                            _spreadsheetService.CreateTextCell($"B{rowIndex}", entry.Value)
                        );
                        sheetData.Append(dataRow);
                        rowIndex++;
                    }
                }

                _logger.LogInformation($"VOID list created at: {voidListFile}");
                return voidListFile;
            }
            catch (Exception ex)
            { 
                _logger.LogError(ex, "Error creating VOID list");
                return string.Empty;
            }
        }
    }
}
