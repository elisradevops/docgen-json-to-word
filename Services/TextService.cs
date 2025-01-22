using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using System;

namespace JsonToWord.Services
{
    public class TextService : ITextService
    {
        private readonly IParagraphService _paragraphService;
        private readonly IRunService _runService;
        private readonly ContentControlService _contentControlService;


        public TextService(IParagraphService paragraphService, IRunService runService)
        {
            _paragraphService = paragraphService;
            _runService = runService;
            _contentControlService = new ContentControlService();
        }
        public void Write(WordprocessingDocument document, string contentControlTitle, WordParagraph wordParagraph)
        {
            var paragraph = _paragraphService.CreateParagraph(wordParagraph);

            if (wordParagraph.Runs != null)
            {
                foreach (var wordRun in wordParagraph.Runs)
                {
                    var run = _runService.CreateRun(wordRun, document);

                    if (!string.IsNullOrEmpty(wordRun.TextStyling.Uri))
                    {
                        try
                        {
                        var id = HyperlinkService.AddHyperlinkRelationship(document.MainDocumentPart, new Uri(wordRun.TextStyling.Uri));
                        var hyperlink = HyperlinkService.CreateHyperlink(id);
                        hyperlink.AppendChild(run);

                        paragraph.AppendChild(hyperlink);
                        }
                        catch (UriFormatException e)
                        {
                            Console.WriteLine(wordRun.TextStyling.Uri+ " is an invalid uri \n" + e.Message);
                            paragraph.AppendChild(run);
                        }
                    }
                   
                    else
                    {
                        paragraph.AppendChild(run);
                    }
                }
            }

            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(paragraph);

            sdtBlock.AppendChild(sdtContentBlock);
        }
    }
}