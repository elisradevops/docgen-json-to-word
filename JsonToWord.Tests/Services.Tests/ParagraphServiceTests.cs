using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services;

namespace JsonToWord.Services.Tests
{
    public class ParagraphServiceTests
    {
        [Fact]
        public void CreateParagraph_WithHeadingZero_ReturnsPlainParagraph()
        {
            var service = new ParagraphService();
            var paragraph = service.CreateParagraph(new WordParagraph { HeadingLevel = 0 }, true);

            Assert.Null(paragraph.ParagraphProperties);
        }

        [Fact]
        public void CreateParagraph_StandardHeading_SetsHeadingStyle()
        {
            var service = new ParagraphService();
            var paragraph = service.CreateParagraph(new WordParagraph { HeadingLevel = 2 }, true);

            var props = paragraph.ParagraphProperties;
            Assert.NotNull(props);
            Assert.Equal("Heading2", props.ParagraphStyleId.Val);
            Assert.Empty(props.Elements<PageBreakBefore>());
        }

        [Fact]
        public void CreateParagraph_CustomHeading_AddsNoPageBreakBefore()
        {
            var service = new ParagraphService();
            var paragraph = service.CreateParagraph(new WordParagraph { HeadingLevel = 2 }, false);

            var props = paragraph.ParagraphProperties;
            Assert.NotNull(props);
            Assert.Equal("Heading1", props.ParagraphStyleId.Val);
            var pageBreak = props.Elements<PageBreakBefore>().FirstOrDefault();
            Assert.NotNull(pageBreak);
            Assert.False(pageBreak.Val.Value);
        }

        [Fact]
        public void ApplyTightSpacing_WithImages_UsesAutoSpacing()
        {
            var service = new ParagraphService();
            var paragraph = new Paragraph(new Run(new Drawing()));

            service.ApplyTightSpacing(paragraph);

            var spacing = paragraph.ParagraphProperties.SpacingBetweenLines;
            Assert.NotNull(spacing);
            Assert.Equal(LineSpacingRuleValues.Auto, spacing.LineRule.Value);
            Assert.Equal("240", spacing.Line);
            Assert.Equal("0", spacing.Before);
            Assert.Equal("60", spacing.After);
        }

        [Fact]
        public void ApplyTightSpacing_WithTextOnly_UsesExactSpacing()
        {
            var service = new ParagraphService();
            var paragraph = new Paragraph(new Run(new Text("text")));

            service.ApplyTightSpacing(paragraph);

            var spacing = paragraph.ParagraphProperties.SpacingBetweenLines;
            Assert.NotNull(spacing);
            Assert.Equal(LineSpacingRuleValues.Exact, spacing.LineRule.Value);
            Assert.Equal("240", spacing.Line);
            Assert.Equal("0", spacing.Before);
            Assert.Equal("0", spacing.After);
        }
    }
}
