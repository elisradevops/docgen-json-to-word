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
        public string Location { get; set; } = string.Empty;
    }

    public class VoidListService : IVoidListService
    {
        private readonly ILogger<VoidListService> _logger;
        private readonly ISpreadsheetService _spreadsheetService;
        private static readonly Regex vlRegex = new Regex(@"#VL-[^#]+#", RegexOptions.IgnoreCase);
        private static readonly Regex validVlRegex = new Regex(@"#VL-\d+(?:\.\d+)?(\s[^#]*)?#", RegexOptions.IgnoreCase);

        public VoidListService(ILogger<VoidListService> logger, ISpreadsheetService spreadsheetService)
        {
            _logger = logger;
            _spreadsheetService = spreadsheetService;
        }

        // Determine if the paragraph is a heading (Heading 1..9)
        private static bool IsHeadingParagraph(Paragraph p)
        {
            // Consider outline level as heading as well (supports custom styles assigned an outline level)
            if (p.ParagraphProperties?.OutlineLevel != null)
                return true;

            var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrEmpty(styleId)) return false;
            var normalized = styleId.Replace(" ", string.Empty);
            return normalized.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
        }

        // Extracts plain text from a paragraph
        private static string GetParagraphText(Paragraph p)
        {
            return p.InnerText?.Trim() ?? string.Empty;
        }

        // Extract only the test case id from a heading text.
        // Expected formats include: "{name} - {id}" where id is numeric, possibly other numbers appear earlier
        private static string ExtractTestCaseIdFromHeading(string heading)
        {
            if (string.IsNullOrWhiteSpace(heading)) return string.Empty;
            // Prefer a number at the end of the string
            var endMatch = Regex.Match(heading, @"(\d+)\s*$");
            if (endMatch.Success) return endMatch.Groups[1].Value;
            // Otherwise, look for a hyphen-delimited id anywhere
            var hyphenMatch = Regex.Match(heading, @"-\s*(\d+)");
            if (hyphenMatch.Success) return hyphenMatch.Groups[1].Value;
            // Fallback: last number in the string
            var allNums = Regex.Matches(heading, @"\d+");
            if (allNums.Count > 0) return allNums[allNums.Count - 1].Value;
            return string.Empty;
        }

        private class RunInfo
        {
            public Word.Run Run { get; set; } = null!;
            public string Text { get; set; } = string.Empty;
            public int StartIndex { get; set; }
        }

        private class ValueAggregate
        {
            public VoidListEntry Entry { get; set; } = null!;
            public HashSet<string> Locations { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private static List<RunInfo> BuildRunInfos(Paragraph paragraph)
        {
            var runInfos = new List<RunInfo>();
            int position = 0;

            foreach (var run in paragraph.Elements<Word.Run>())
            {
                string runText = run.InnerText ?? string.Empty;
                runInfos.Add(new RunInfo
                {
                    Run = run,
                    Text = runText,
                    StartIndex = position
                });
                position += runText.Length;
            }

            return runInfos;
        }

        private static int GetRunIndexAtPosition(List<RunInfo> runInfos, int position)
        {
            for (int i = 0; i < runInfos.Count; i++)
            {
                int start = runInfos[i].StartIndex;
                int end = start + runInfos[i].Text.Length;
                if (position < end)
                {
                    return i;
                }
            }

            return runInfos.Count - 1;
        }

        private static void RewriteRunsForMatch(Paragraph paragraph, List<RunInfo> runInfos, Match match, bool isValidCode, string key, string originalMatchValue)
        {
            if (runInfos.Count == 0)
            {
                return;
            }

            int startIndex = GetRunIndexAtPosition(runInfos, match.Index);
            int endIndex = GetRunIndexAtPosition(runInfos, match.Index + match.Length - 1);

            if (startIndex < 0 || endIndex < 0 || startIndex >= runInfos.Count || endIndex >= runInfos.Count)
            {
                return;
            }

            var startRunInfo = runInfos[startIndex];
            var endRunInfo = runInfos[endIndex];

            int startOffset = match.Index - startRunInfo.StartIndex;
            int endOffset = match.Index + match.Length - endRunInfo.StartIndex;

            string prefix = startOffset > 0 ? startRunInfo.Text.Substring(0, startOffset) : string.Empty;
            string suffix = endOffset < endRunInfo.Text.Length ? endRunInfo.Text.Substring(endOffset) : string.Empty;

            var runsToRemove = runInfos
                .Skip(startIndex)
                .Take(endIndex - startIndex + 1)
                .Select(info => info.Run)
                .Distinct()
                .ToList();

            var anchorRun = startRunInfo.Run;
            var newRuns = new List<Word.Run>();

            if (!string.IsNullOrEmpty(prefix))
            {
                newRuns.Add(CreateRunFromReference(anchorRun, prefix));
            }

            Word.Run matchRun = isValidCode
                ? CreateFormattedMatchRun(anchorRun, "#" + key, "0000FF")
                : CreateFormattedMatchRun(anchorRun, originalMatchValue, "FF0000");

            newRuns.Add(matchRun);

            if (!string.IsNullOrEmpty(suffix))
            {
                newRuns.Add(CreateRunFromReference(endRunInfo.Run, suffix));
            }

            foreach (var newRun in newRuns)
            {
                paragraph.InsertBefore(newRun, anchorRun);
            }

            foreach (var run in runsToRemove)
            {
                run.Remove();
            }
        }

        private static Word.Run CreateRunFromReference(Word.Run referenceRun, string text)
        {
            var run = new Word.Run();

            if (referenceRun?.RunProperties != null)
            {
                run.RunProperties = (Word.RunProperties)referenceRun.RunProperties.CloneNode(true);
            }

            run.AppendChild(CreateTextElement(text));
            return run;
        }

        private static Word.Run CreateFormattedMatchRun(Word.Run referenceRun, string text, string colorHex)
        {
            var run = new Word.Run();

            Word.RunProperties runProperties;
            if (referenceRun?.RunProperties != null)
            {
                runProperties = (Word.RunProperties)referenceRun.RunProperties.CloneNode(true);
                runProperties.RemoveAllChildren<Word.Color>();
                runProperties.RemoveAllChildren<Word.Underline>();
                runProperties.RemoveAllChildren<Word.Bold>();
            }
            else
            {
                runProperties = new Word.RunProperties();
            }

            runProperties.AppendChild(new Word.Bold());
            runProperties.AppendChild(new Word.Color { Val = colorHex });
            runProperties.AppendChild(new Word.Underline { Val = Word.UnderlineValues.Single });
            run.RunProperties = runProperties;

            run.AppendChild(CreateTextElement(text));
            return run;
        }

        private static Word.Text CreateTextElement(string text)
        {
            return new Word.Text(text) { Space = SpaceProcessingModeValues.Preserve };
        }

        public List<string> CreateVoidList(string docPath)
        {
            List<string> filesToZip = new List<string>();
            string docName = Path.GetFileName(docPath)?.Replace(':', '_');
            string voidListFile = Path.Combine(Path.GetDirectoryName(docPath) ?? string.Empty, docName + " - VOID LIST.xlsx");
            
            var allMatches = new List<VoidListEntry>();
            var validationErrors = new List<string>();
            var entriesByKey = new Dictionary<string, Dictionary<string, ValueAggregate>>(StringComparer.OrdinalIgnoreCase);

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

                    string currentLocationId = string.Empty;
                    foreach (var p in mainPart.Document.Body.Descendants<Paragraph>().ToList())
                    {
                        // Update current location when encountering a heading paragraph
                        if (IsHeadingParagraph(p))
                        {
                            var headingText = GetParagraphText(p);
                            currentLocationId = ExtractTestCaseIdFromHeading(headingText);
                        }

                        int searchIndex = 0;
                        while (true)
                        {
                            var runInfos = BuildRunInfos(p);
                            if (runInfos.Count == 0)
                            {
                                break;
                            }

                            string paragraphText = string.Concat(runInfos.Select(info => info.Text));
                            if (string.IsNullOrEmpty(paragraphText) || searchIndex >= paragraphText.Length)
                            {
                                break;
                            }

                            Match match = vlRegex.Match(paragraphText, searchIndex);
                            if (!match.Success)
                            {
                                break;
                            }

                            string originalMatchValue = match.Value;
                            bool isValidCode = validVlRegex.IsMatch(originalMatchValue);
                            string[] parts = originalMatchValue.Trim('#').Split(new[] { ' ' }, 2);
                            string key = parts[0].ToUpper();
                            string value = parts.Length > 1 ? parts[1].Trim() : string.Empty;

                            if (isValidCode)
                            {
                                if (!entriesByKey.TryGetValue(key, out var valuesForKey))
                                {
                                    valuesForKey = new Dictionary<string, ValueAggregate>(StringComparer.OrdinalIgnoreCase);
                                    entriesByKey[key] = valuesForKey;
                                }

                                string normalizedValue = value;
                                string normalizedLocation = string.IsNullOrWhiteSpace(currentLocationId) ? "Unknown" : currentLocationId;

                                if (!valuesForKey.TryGetValue(normalizedValue, out var aggregate))
                                {
                                    int currentIndex = valuesForKey.Count + 1;
                                    var entry = new VoidListEntry
                                    {
                                        Key = key,
                                        Value = value,
                                        Index = currentIndex,
                                        DisplayKey = currentIndex > 1 ? $"{key}-{currentIndex}" : key,
                                        Location = normalizedLocation
                                    };

                                    aggregate = new ValueAggregate
                                    {
                                        Entry = entry
                                    };
                                    aggregate.Locations.Add(normalizedLocation);
                                    valuesForKey[normalizedValue] = aggregate;
                                    allMatches.Add(entry);
                                }
                                else
                                {
                                    aggregate.Locations.Add(normalizedLocation);
                                }
                            }
                            else
                            {
                                var locDisplay = string.IsNullOrWhiteSpace(currentLocationId) ? "Unknown" : currentLocationId;
                                validationErrors.Add($"Invalid VL code found: {originalMatchValue} at Work Item: {locDisplay}. Expected format: #VL-<number> [text]# (e.g., #VL-123 Description#).");
                            }

                            RewriteRunsForMatch(p, runInfos, match, isValidCode, key, originalMatchValue);

                            searchIndex = Math.Max(match.Index + 1, 0);
                        }
                    }

                    mainPart.Document.Save();
                }

                // Add validation errors for conflicting values per key
                foreach (var keyEntries in entriesByKey)
                {
                    if (keyEntries.Value.Count <= 1)
                    {
                        continue;
                    }

                    var valueSummaries = keyEntries.Value.Select(pair =>
                    {
                        string normalizedValue = pair.Key;
                        string displayValue = string.IsNullOrWhiteSpace(normalizedValue) ? "<empty>" : pair.Value.Entry.Value;
                        IEnumerable<string> locations = pair.Value.Locations.Count > 0 ? pair.Value.Locations : new[] { "Unknown" };
                        return $"\"{displayValue}\" @ [{string.Join(", ", locations)}]";
                    });

                    validationErrors.Add($"Conflicting VL content found: {keyEntries.Key} appears with multiple descriptions {string.Join("; ", valueSummaries)}.");
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
                        if (decimal.TryParse(keyPart, out decimal numericKey))
                            return numericKey;
                        return decimal.MaxValue; // Put non-numeric keys at the end
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
