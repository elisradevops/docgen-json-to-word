using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using System;
using System.Collections.Generic;

namespace JsonToWord.Services
{
    public class TextService : ITextService
    {
        private readonly IContentControlService _contentControlService;
        private readonly IParagraphService _paragraphService;
        private readonly IRunService _runService;
        private int _bookmarkIdCounter;
        private readonly HashSet<string> _emittedBookmarks = new HashSet<string>();


        public TextService(IContentControlService contentControlService, IParagraphService paragraphService, IRunService runService)
        {
            _paragraphService = paragraphService;
            _runService = runService;
            _contentControlService = contentControlService;
        }
        public void Write(WordprocessingDocument document, string contentControlTitle, WordParagraph wordParagraph, bool isUnderStandardHeading)
        {
            var paragraph = _paragraphService.CreateParagraph(wordParagraph, isUnderStandardHeading);

            // Emit bookmark if BookmarkName is set and not yet emitted (dedupe)
            BookmarkEnd bookmarkEnd = null;
            if (!string.IsNullOrEmpty(wordParagraph.BookmarkName) && _emittedBookmarks.Add(wordParagraph.BookmarkName))
            {
                var bmId = (++_bookmarkIdCounter).ToString();
                paragraph.AppendChild(new BookmarkStart { Id = bmId, Name = wordParagraph.BookmarkName });
                bookmarkEnd = new BookmarkEnd { Id = bmId };
            }

            if (wordParagraph.Runs != null)
            {
                foreach (var wordRun in wordParagraph.Runs)
                {
                    var run = _runService.CreateRun(wordRun);

                    if (!string.IsNullOrEmpty(wordRun.Uri))
                    {
                        try
                        {
                        var id = HyperlinkService.AddHyperlinkRelationship(document.MainDocumentPart, new Uri(wordRun.Uri));
                        var hyperlink = HyperlinkService.CreateHyperlink(id);
                        hyperlink.AppendChild(run);

                        paragraph.AppendChild(hyperlink);
                        }
                        catch (UriFormatException e)
                        {
                            Console.WriteLine(wordRun.Uri+ " is an invalid uri \n" + e.Message);
                            paragraph.AppendChild(run);
                        }
                    }

                    else
                    {
                        paragraph.AppendChild(run);
                    }
                }
            }

            // Close bookmark after all runs
            if (bookmarkEnd != null)
            {
                paragraph.AppendChild(bookmarkEnd);
            }

            if (_contentControlService.WriteParagraphToContentControl(document, contentControlTitle, paragraph))
                return;

            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);
            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(paragraph);
            sdtBlock.AppendChild(sdtContentBlock);
        }
    }
}
