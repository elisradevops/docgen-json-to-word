using System;
using Amazon.Runtime.Internal.Util;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using Microsoft.Extensions.Logging;

namespace JsonToWord.Services
{
    internal class RunService: IRunService
    {
        private readonly IPictureService _pictureService;
        private ILogger<RunService> _logger;
        public RunService(IPictureService pictureService, ILogger<RunService> logger)
        {
            _pictureService = pictureService;
            _logger = logger;
        }

        public Run CreateRun(WordRun wordRun, WordprocessingDocument document)
        {
            var run = new Run();
            var runProperties = new RunProperties();
            if (wordRun.Type == "text")
            {
                SetHyperlink(wordRun, runProperties);
                SetBold(wordRun, runProperties);
                SetItalic(wordRun, runProperties);
                SetUnderline(wordRun, runProperties);
                SetSize(wordRun, runProperties);
                SetColor(wordRun, runProperties);
                run.AppendChild(runProperties);
                SetBreak(wordRun, run);
                SetText(wordRun, run);
            }
            if(wordRun.Type == "break")
            {
                run.AppendChild(new Break());
            }
            else if(wordRun.Type == "image")
            {
                if (!string.IsNullOrEmpty(wordRun.Src))
                {
                    var drawing = _pictureService.CreateDrawing(document.MainDocumentPart, wordRun.Src);
                    run.AppendChild(drawing);
                }
            }
           
            return run;
        }

        private void SetColor(WordRun wordRun, RunProperties runProperties)
        {
            if (!string.IsNullOrEmpty(wordRun.TextStyling.FontColor))
            {
                try
                {
                    System.Drawing.Color color = System.Drawing.Color.FromName(wordRun.TextStyling.FontColor);
                    string colorHex = color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
                    Color wordColor = new Color() { Val = colorHex };
                    runProperties.AppendChild(wordColor);
                }
                catch (Exception exception)
                {
                   _logger.LogError(exception, "Invalid color value");
                }
            }
        }

        private static void SetBreak(WordRun wordRun, Run run)
        {
            if (wordRun.TextStyling.InsertLineBreak)
                run.AppendChild(new Break());
        }

        private static void SetText(WordRun wordRun, Run run)
        {
            if (string.IsNullOrEmpty(wordRun.Value))
                return;

            var text = new Text { Text = wordRun.Value };

            if (wordRun.TextStyling.InsertSpace)
                text.Space = SpaceProcessingModeValues.Preserve;

            run.AppendChild(text);
        }

        private static void SetSize(WordRun wordRun, RunProperties runProperties)
        {
            if (wordRun.TextStyling.Size != 0)
                runProperties.FontSize = new FontSize { Val = new StringValue((wordRun.TextStyling.Size * 2).ToString()) };
        }

        private static void SetUnderline(WordRun wordRun, RunProperties runProperties)
        {
            if (wordRun.TextStyling.Underline && string.IsNullOrEmpty(wordRun.TextStyling.Uri))
                AddUnderline(runProperties);
        }

        private static void SetItalic(WordRun wordRun, RunProperties runProperties)
        {
            if (!wordRun.TextStyling.Italic)
                return;

            var italic = new Italic();
            var italicComplexScript = new ItalicComplexScript();

            runProperties.AppendChild(italic);
            runProperties.AppendChild(italicComplexScript);
        }

        private static void SetBold(WordRun wordRun, RunProperties runProperties)
        {
            if (!wordRun.TextStyling.Bold)
                return;

            var bold = new Bold();
            var boldComplexScript = new BoldComplexScript();

            runProperties.AppendChild(bold);
            runProperties.AppendChild(boldComplexScript);
        }

        private static void SetHyperlink(WordRun wordRun, RunProperties runProperties)
        {
            if (!string.IsNullOrEmpty(wordRun.TextStyling.Uri))
            {
                var runStyle = new RunStyle() { Val = "Hyperlink" };
                var color = new Color() { Val = "auto", ThemeColor = ThemeColorValues.Hyperlink };

                runProperties.AppendChild(runStyle);
                runProperties.AppendChild(color);
                AddUnderline(runProperties);
            }
            else
            {
                var runFonts = new RunFonts { Ascii = wordRun.TextStyling.Font, HighAnsi = wordRun.TextStyling.Font, ComplexScript = wordRun.TextStyling.Font };
                runProperties.AppendChild(runFonts);
            }
        }

        private static void AddUnderline(RunProperties runProperties)
        {
            var underline = new Underline() { Val = UnderlineValues.Single };
            runProperties.AppendChild(underline);
        }
    }
}