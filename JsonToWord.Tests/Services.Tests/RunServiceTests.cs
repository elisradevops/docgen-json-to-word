using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services;
using JsonToWord.Services.Interfaces;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class RunServiceTests
    {
        [Fact]
        public void CreateRun_WithHyperlink_SetsHyperlinkStyleAndUnderline()
        {
            var pictureService = new Mock<IPictureService>();
            var service = new RunService(pictureService.Object);
            var wordRun = new WordRun
            {
                Uri = "https://example.com",
                Text = "Link",
                InsertSpace = true,
                Font = "Arial"
            };

            var run = service.CreateRun(wordRun);
            var props = run.GetFirstChild<RunProperties>();

            Assert.NotNull(props);
            Assert.Equal("Hyperlink", props.RunStyle.Val.Value);
            Assert.NotNull(props.Underline);
            Assert.Null(props.RunFonts);

            var text = run.Descendants<Text>().FirstOrDefault();
            Assert.NotNull(text);
            Assert.Equal(SpaceProcessingModeValues.Preserve, text.Space.Value);
        }

        [Fact]
        public void CreateRun_WithFormatting_SetsPropertiesAndBreaks()
        {
            var pictureService = new Mock<IPictureService>();
            var service = new RunService(pictureService.Object);
            var wordRun = new WordRun
            {
                Bold = true,
                Italic = true,
                Underline = true,
                Size = 10,
                Font = "Courier New",
                FontColor = "Red",
                Text = "Hello",
                InsertLineBreak = true
            };

            var run = service.CreateRun(wordRun);
            var props = run.GetFirstChild<RunProperties>();

            Assert.NotNull(props);
            Assert.NotNull(props.Bold);
            Assert.NotNull(props.Italic);
            Assert.NotNull(props.Underline);
            Assert.Equal("20", props.FontSize.Val.Value);
            Assert.Equal("Courier New", props.RunFonts.Ascii.Value);
            Assert.Equal("FF0000", props.Color.Val.Value);
            Assert.True(run.Descendants<Break>().Any(b => b.Type == null));
        }

        [Fact]
        public void CreateRun_WithPageBreak_AddsPageBreak()
        {
            var pictureService = new Mock<IPictureService>();
            var service = new RunService(pictureService.Object);
            var wordRun = new WordRun
            {
                Text = "Page Break",
                InsertPageBreak = true
            };

            var run = service.CreateRun(wordRun);

            Assert.True(run.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page));
        }
    }
}
