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
    public class VoidListEntry
    {
        public string Key { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public int Index { get; set; }
        public string DisplayKey { get; set; } = string.Empty;
    }

    public class VoidListService : IVoidListService
    {
        private readonly ILogger<VoidListService> _logger;
        private readonly ISpreadsheetService _spreadsheetService;
        private static readonly Regex vlRegex = new Regex(@"#VL-[^#]+#", RegexOptions.IgnoreCase);
        private static readonly Regex validVlRegex = new Regex(@"#VL-\d+[^#]*#", RegexOptions.IgnoreCase);

        public VoidListService(ILogger<VoidListService> logger, ISpreadsheetService spreadsheetService)
        {
            _logger = logger;
            _spreadsheetService = spreadsheetService;
        }

        public List<string> CreateVoidList(string docPath)
        {
            List<string> filesToZip = new List<string>();
            string docName = Path.GetFileName(docPath)?.Replace(':', '_');
            string voidListFile = Path.Combine(Path.GetDirectoryName(docPath) ?? string.Empty, docName + " - VOID LIST.xlsx");
            
            var allMatches = new List<VoidListEntry>();
            var validationErrors = new List<string>();
            var duplicateTracker = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docPath, true))
                {
                    var mainPart = wordDoc.MainDocumentPart;
                    if (mainPart?.Document?.Body == null)
                    {
                        _logger.LogWarning("Document body is null. Cannot process for VOID list.");
                        return filesToZip;
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
                                bool isValidCode = validVlRegex.IsMatch(originalMatchValue);
                                
                                // In-document replacement logic
                                string[] parts = originalMatchValue.Trim('#').Split(new[] { ' ' }, 2);
                                string key = parts[0].ToUpper();
                                string value = parts.Length > 1 ? parts[1] : string.Empty;
                                
                                // Track duplicates and validation
                                if (isValidCode)
                                {
                                    if (duplicateTracker.ContainsKey(key))
                                    {
                                        duplicateTracker[key]++;
                                    }
                                    else
                                    {
                                        duplicateTracker[key] = 1;
                                    }
                                    
                                    // Add to matches list with index for duplicates
                                    int currentIndex = duplicateTracker[key];
                                    allMatches.Add(new VoidListEntry 
                                    { 
                                        Key = key, 
                                        Value = value, 
                                        Index = currentIndex,
                                        DisplayKey = currentIndex > 1 ? $"{key}-{currentIndex}" : key
                                    });
                                    
                                    string newMatchValue = "#" + key;
                                    var matchRun = new Word.Run(new Word.Text(newMatchValue) { Space = SpaceProcessingModeValues.Preserve });
                                    Word.RunProperties rp = (run.RunProperties != null) ? (Word.RunProperties)run.RunProperties.CloneNode(true) : new Word.RunProperties();
                                    rp.Append(new Word.Bold());
                                    rp.Append(new Word.Color() { Val = "0000FF" });
                                    rp.Append(new Word.Underline() { Val = Word.UnderlineValues.Single });
                                    matchRun.RunProperties = rp;
                                    newRuns.Add(matchRun);
                                }
                                else
                                {
                                    // Invalid code - mark as red and add to validation errors
                                    validationErrors.Add($"Invalid VL code found: {originalMatchValue} - Code must be in format #VL-[NUMBER]..#");
                                    
                                    var matchRun = new Word.Run(new Word.Text(originalMatchValue) { Space = SpaceProcessingModeValues.Preserve });
                                    Word.RunProperties rp = (run.RunProperties != null) ? (Word.RunProperties)run.RunProperties.CloneNode(true) : new Word.RunProperties();
                                    rp.Append(new Word.Bold());
                                    rp.Append(new Word.Color() { Val = "FF0000" }); // Red color for invalid codes
                                    rp.Append(new Word.Underline() { Val = Word.UnderlineValues.Single });
                                    matchRun.RunProperties = rp;
                                    newRuns.Add(matchRun);
                                }

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

                // Add duplicate validation errors
                foreach (var duplicate in duplicateTracker.Where(d => d.Value > 1))
                {
                    validationErrors.Add($"Duplicate VL code found: {duplicate.Key} appears {duplicate.Value} times");
                }
                
                // Create validation report
                if (validationErrors.Count > 0)
                {
                    string validationReportFile = Path.Combine(Path.GetDirectoryName(docPath) ?? string.Empty, docName + " - VALIDATION REPORT.txt");
                    File.WriteAllLines(validationReportFile, new[] { "VOID List Validation Report", "=" + new string('=', 28), "" }.Concat(validationErrors));
                    filesToZip.Add(validationReportFile);
                    _logger.LogInformation($"Validation report created at: {validationReportFile}");
                }

                if (allMatches.Count == 0)
                {
                    _logger.LogInformation("No valid VOID list matches found in the document.");
                    return filesToZip;
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

                    // Sort entries by key (natural numeric order for VL codes)
                    var sortedMatches = allMatches.OrderBy(entry => 
                    {
                        // Extract numeric part from VL-X format for proper sorting
                        var keyPart = entry.Key.Replace("VL-", "");
                        if (int.TryParse(keyPart, out int numericKey))
                            return numericKey;
                        return int.MaxValue; // Put non-numeric keys at the end
                    }).ThenBy(entry => entry.Index).ToList();

                    // Add data rows
                    uint rowIndex = 2;
                    foreach (var entry in sortedMatches)
                    {
                        Row dataRow = new Row() { RowIndex = rowIndex };
                        dataRow.Append(
                            _spreadsheetService.CreateTextCell($"A{rowIndex}", entry.DisplayKey),
                            _spreadsheetService.CreateTextCell($"B{rowIndex}", entry.Value)
                        );
                        sheetData.Append(dataRow);
                        rowIndex++;
                    }
                }

                _logger.LogInformation($"VOID list created at: {voidListFile}");
                filesToZip.Add(voidListFile);
                return filesToZip;
            }
            catch (Exception ex)
            { 
                _logger.LogError(ex, "Error creating VOID list");
                return filesToZip;
            }
        }
    }
}
