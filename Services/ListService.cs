using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JsonToWord.Services
{
    public class ListService : IListService
    {
        private readonly IParagraphService _paragraphService;
        private readonly IRunService _runService;
        private readonly ILogger<ListService> _logger;
        private readonly ContentControlService _contentControlService;

        public ListService(
            ILogger<ListService> logger,
            IParagraphService paragraphService,
            IRunService runService)
        {
            _paragraphService = paragraphService;
            _runService = runService;
            _logger = logger;
            _contentControlService = new ContentControlService();
        }

        public void Insert(WordprocessingDocument document, string contentControlTitle, WordList wordList)
        {
            // 1) Validate input
            if (!IsValidWordList(wordList))
            {
                _logger.LogWarning("List is empty or invalid");
                return;
            }

            // 2) Find the target content control
            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);
            if (sdtBlock == null)
            {
                _logger.LogWarning($"Could not find content control '{contentControlTitle}'.");
                return;
            }

            // 3) Ensure we have a NumberingDefinitionsPart
            var numberingPart = EnsureNumberingPart(document);

            // Decide single-level vs multi-level
            bool multiLevel = (wordList.ListItems.Count > 1);

            // 5) Create the numbering definition
            int numId = CreateNumberingDefinitionWithNsid(
                numberingPart,
                wordList.IsOrdered,
                multiLevel);

            // 6) Build an SdtContentBlock
            var sdtContentBlock = new SdtContentBlock();

            // 7) For each list item, create a paragraph
            foreach (var item in wordList.ListItems)
            {
                // item.Level = 0,1,2,... 
                // If single-level => we clamp to 0 in InitParagraphForListItem.
                var para = _paragraphService.InitParagraphForListItem(item, wordList.IsOrdered, numId, multiLevel);
                AppendRunsToParagraph(para, item.Runs, document);
                sdtContentBlock.AppendChild(para);
            }

            // 8) Append to the content control
            sdtBlock.AppendChild(sdtContentBlock);

            // 9) Optionally add a blank paragraph to separate from next content
            var emptyPara = new Paragraph(
                new ParagraphProperties(),
                new Run(new Text(""))
            );
            sdtBlock.AppendChild(emptyPara);

            // Save
            document.MainDocumentPart.Document.Save();
        }


        /// <summary>
        /// Creates a new AbstractNum + NumberingInstance with a unique nsid + tmpl.
        /// This prevents Word from merging them if they look "identical".
        /// </summary>
        private int CreateNumberingDefinitionWithNsid(
            NumberingDefinitionsPart numberingPart,
            bool isOrdered,
            bool multiLevel)
        {

            // 1) fetch the next IDs
            int newAbsId = GetNextAbstractNumId(numberingPart);
            int newNumId = GetNextNumId(numberingPart);

            // A new AbstractNum with the same ID as numId
            var abstractNum = new AbstractNum { AbstractNumberId = newAbsId };

            // (A) Add a unique nsid (8 hex chars) + a random <w:tmpl> code
            //     You can do real random or a short GUID
            abstractNum.AppendChild(new Nsid { Val = RandomHex(8) });

            // (B) multi-level or single-level
            if (multiLevel)
                abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
            else
                abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.SingleLevel });

            abstractNum.AppendChild(new TemplateCode { Val = RandomHex(8) });

            // We'll define up to 9 levels if multi-level, else just 1
            int maxLevels = multiLevel ? 9 : 1;

            // Some rotation sets
            string[] bulletSymbols = { "·", "o", "§" };
            NumberFormatValues[] numberFormats =
            {
                NumberFormatValues.Decimal,
                NumberFormatValues.LowerLetter,
                NumberFormatValues.LowerRoman
            };

            for (int lvl = 0; lvl < maxLevels; lvl++)
            {
                var level = new Level { LevelIndex = lvl };

                // <w:start w:val="1"/>
                var start = new StartNumberingValue { Val = 1 };
                level.AppendChild(start);

                if (isOrdered)
                {
                    // Rotate decimal, lowerLetter, lowerRoman
                    var chosenFmt = numberFormats[lvl % numberFormats.Length];
                    level.AppendChild(new NumberingFormat { Val = chosenFmt });

                    // Multi-level pattern: "%1.", "%1.%2.", etc.
                    string lvlTextVal = BuildMultiLevelText(lvl + 1);
                    level.AppendChild(new LevelText { Val = lvlTextVal });
                }
                else
                {
                    // Bulleted
                    level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
                    var symbol = bulletSymbols[lvl % bulletSymbols.Length];
                    level.AppendChild(new LevelText { Val = symbol });
                }

                // <w:lvlJc w:val="left"/>
                level.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });

                // Indentation
                var prevPPr = new PreviousParagraphProperties(
                    new Indentation
                    {
                        Left = (720 * (lvl + 1)).ToString(),
                        Hanging = "360"
                    }
                );
                level.AppendChild(prevPPr);

                abstractNum.Append(level);
            }


            // 3) Insert the new <w:abstractNum> 
            //    right after the last <w:abstractNum> but before <w:num>
            InsertAbstractNumAfterExistingAbstractNum(numberingPart.Numbering, abstractNum);


            // 4) Now create <w:num> referencing it
            var numberingInstance = new NumberingInstance { NumberID = newNumId };
            numberingInstance.AppendChild(new AbstractNumId { Val = newAbsId });

            // (optional) level override so we start at 1
            var lvlOverride = new LevelOverride { LevelIndex = 0 };
            lvlOverride.AppendChild(new StartOverrideNumberingValue { Val = 1 });
            numberingInstance.AppendChild(lvlOverride);

            InsertNumberingInstanceAfterExistingNums(numberingPart.Numbering, numberingInstance);

            return newNumId;
        }

        private void InsertAbstractNumAfterExistingAbstractNum(Numbering numberingRoot, AbstractNum newAbs)
        {
            // find last <w:abstractNum> in the numbering
            var lastAbs = numberingRoot.Elements<AbstractNum>().LastOrDefault();
            if (lastAbs != null)
            {
                numberingRoot.InsertAfter(newAbs, lastAbs);
            }
            else
            {
                // if none, just put it at start
                numberingRoot.PrependChild(newAbs);
            }
        }

        private void InsertNumberingInstanceAfterExistingNums(Numbering numberingRoot, NumberingInstance newNum)
        {
            // find last <w:num> in the numbering
            var lastNum = numberingRoot.Elements<NumberingInstance>().LastOrDefault();
            if (lastNum != null)
            {
                numberingRoot.InsertAfter(newNum, lastNum);
            }
            else
            {
                // if no <w:num> exist, we insert after the last <w:abstractNum> or start
                // simplest approach: append at the end
                numberingRoot.AppendChild(newNum);
            }
        }


        private string RandomHex(int length)
        {
            // E.g. produce a short random hex string
            // This is just an example; you could use a GUID substring, etc.
            var rng = new Random();
            var sb = new StringBuilder(length);
            for (int i = 0; i < length; i++)
            {
                sb.Append(rng.Next(16).ToString("X")); // uppercase hex
            }
            return sb.ToString();
        }


        private int GetNextAbstractNumId(NumberingDefinitionsPart numberingPart)
        {
            // Among all <w:abstractNum w:abstractNumId="X">
            var existing = numberingPart.Numbering
                .Elements<AbstractNum>()
                .Select(a => a.AbstractNumberId.Value)
                .DefaultIfEmpty(0)  // in case no <w:abstractNum> exist
                .Max();

            return existing + 1;
        }

        private int GetNextNumId(NumberingDefinitionsPart numberingPart)
        {
            // Among all <w:num w:numId="X">
            var existing = numberingPart.Numbering
                .Elements<NumberingInstance>()
                .Select(n => n.NumberID.Value)
                .DefaultIfEmpty(0)  // in case no <w:num> exist
                .Max();

            return existing + 1;
        }


        private bool IsValidWordList(WordList wordList)
        {
            return wordList != null
                   && wordList.ListItems != null
                   && wordList.ListItems.Count > 0;
        }

        private NumberingDefinitionsPart EnsureNumberingPart(WordprocessingDocument document)
        {
            var numberingPart = document.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
            }
            return numberingPart;
        }

        /// <summary>
        /// Finds the next unique integer for both <w:abstractNum w:abstractNumId="...">
        /// and <w:num w:numId="..."> so we don't overwrite existing ones or heading numbering.
        /// </summary>
        private int GetNextUniqueId(NumberingDefinitionsPart numberingPart)
        {
            var existingNumIds = numberingPart.Numbering
                .Elements<NumberingInstance>()
                .Select(n => n.NumberID.Value);

            var existingAbsIds = numberingPart.Numbering
                .Elements<AbstractNum>()
                .Select(a => a.AbstractNumberId.Value);

            var allUsed = existingNumIds.Concat(existingAbsIds).ToList();
            if (!allUsed.Any()) return 1;

            return allUsed.Max() + 1;
        }

        /// <summary>
        /// Creates a new AbstractNum (and corresponding NumberingInstance) for
        /// either a single-level list or a multi-level list, ordered or unordered.
        /// This single definition can be used by all items in the WordList.
        /// 
        /// We'll define up to 9 levels if multi-level. If single-level, just define lvl=0.
        /// We also ensure the numbering starts at 1 again (so each new list doesn't
        /// continue from the previous).
        /// </summary>
        private void CreateNumberingDefinition(
            NumberingDefinitionsPart numberingPart,
            int numId,
            bool isOrdered,
            bool multiLevel)
        {
            // The AbstractNum ID is the same as numId for convenience
            var abstractNum = new AbstractNum { AbstractNumberId = numId };

            // If multi-level, we do HybridMultilevel or Multilevel;
            // If single-level, do SingleLevel.
            if (multiLevel)
                abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
            else
                abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.SingleLevel });

            // We'll define at least 1 level, or up to 9 if multi-level
            int maxLevels = multiLevel ? 9 : 1;

            // Bullet symbols and numeric formats
            string[] bulletSymbols = { "·", "o", "§" };
            NumberFormatValues[] numberFormats =
            {
                NumberFormatValues.Decimal,
                NumberFormatValues.LowerLetter,
                NumberFormatValues.LowerRoman
            };

            for (int lvl = 0; lvl < maxLevels; lvl++)
            {
                var level = new Level { LevelIndex = lvl };

                // Always start from 1
                var start = new StartNumberingValue { Val = 1 };
                level.AppendChild(start);

                // Decide bullet vs. number
                if (isOrdered)
                {
                    // rotate decimal, lowerLetter, lowerRoman
                    var chosenFmt = numberFormats[lvl % numberFormats.Length];
                    level.AppendChild(new NumberingFormat { Val = chosenFmt });

                    // e.g., for multi-level: "1.", "1.%2.", etc. 
                    // We'll do a simplified version (just "1." for all levels)
                    // or we can do the typical multi-level pattern:
                    string levelTextVal = BuildMultiLevelText(lvl + 1);
                    level.AppendChild(new LevelText { Val = levelTextVal });
                }
                else
                {
                    // Bulleted
                    level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
                    var symbol = bulletSymbols[lvl % bulletSymbols.Length];
                    level.AppendChild(new LevelText { Val = symbol });
                }

                // Justify left
                level.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });

                // Indentation
                var prevPPr = new PreviousParagraphProperties(
                    new Indentation
                    {
                        Left = (720 * (lvl + 1)).ToString(),
                        Hanging = "360"
                    }
                );
                level.AppendChild(prevPPr);

                // Append to AbstractNum
                abstractNum.Append(level);

                // If it's single-level, just break after one level
                if (!multiLevel) break;
            }

            // Append the AbstractNum
            numberingPart.Numbering.Append(abstractNum);

            // Create the <w:num> referencing it
            var numberingInstance = new NumberingInstance { NumberID = numId };
            numberingInstance.Append(new AbstractNumId { Val = numId });

            // This ensures the new list starts at 1 again, rather than continuing
            // from a previous list. We do a <w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>
            var lvlOverride = new LevelOverride { LevelIndex = 0 };
            lvlOverride.Append(new StartOverrideNumberingValue { Val = 1 });
            numberingInstance.Append(lvlOverride);

            numberingPart.Numbering.Append(numberingInstance);
        }

        /// <summary>
        /// Typical multi-level text: 
        ///  depth=1 => "%1."
        ///  depth=2 => "%1.%2."
        ///  ...
        /// </summary>
        private string BuildMultiLevelText(int depth)
        {
            var sb = new StringBuilder();
            for (int i = 1; i <= depth; i++)
            {
                sb.Append($"%{i}.");
            }
            return sb.ToString();
        }

        /// <summary>
        /// For each item, append runs (and potentially hyperlinks)
        /// to the paragraph.
        /// </summary>
        private void AppendRunsToParagraph(Paragraph paragraph, IEnumerable<WordRun> runs, WordprocessingDocument document)
        {
            if (runs == null) return;

            foreach (var wordRun in runs)
            {
                var runElement = _runService.CreateRun(wordRun, document);

                if (!string.IsNullOrEmpty(wordRun.TextStyling?.Uri))
                {
                    try
                    {
                        var hyperlinkId = HyperlinkService.AddHyperlinkRelationship(
                            document.MainDocumentPart,
                            new Uri(wordRun.TextStyling.Uri)
                        );
                        var hyperlink = HyperlinkService.CreateHyperlink(hyperlinkId);
                        hyperlink.AppendChild(runElement);

                        paragraph.AppendChild(hyperlink);
                    }
                    catch (UriFormatException e)
                    {
                        Console.WriteLine($"{wordRun.TextStyling.Uri} is invalid: {e.Message}");
                        paragraph.AppendChild(runElement);
                    }
                }
                else
                {
                    paragraph.AppendChild(runElement);
                }
            }
        }
    }
}
