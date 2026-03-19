using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services;
using Microsoft.Extensions.Logging;
using Moq;

namespace JsonToWord.Services.Tests
{
    public class SectionPlaceholderServiceTests
    {
        [Fact]
        public void ResolveSectionPlaceholders_UsesMarkerAnchor_WhenAnchorIsProvided()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            body.Append(CreateHeadingParagraph("System Requirements", "Heading1")); // 1
            body.Append(
                new SdtBlock(
                    new SdtProperties(),
                    new SdtContentBlock(new Paragraph(new Run(new Text("{{section-anchor:requirements-root}}"))))
                )
            );

            body.Append(CreateHeadingParagraph("Qualification Provisions", "Heading1")); // 2
            body.Append(CreateSingleCellTable("{{section:requirements-root:2.3}}"));

            var logger = new Mock<ILogger<SectionPlaceholderService>>();
            var service = new SectionPlaceholderService(logger.Object);

            service.ResolveSectionPlaceholders(document);

            var resolvedText = body.Descendants<Text>().Last().Text;
            Assert.Equal("1.2.3", resolvedText);
            Assert.DoesNotContain("{{section-anchor:requirements-root}}", body.InnerText);
        }

        [Fact]
        public void ResolveSectionPlaceholders_UsesParentHeading_WhenNoAnchorProvided()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            body.Append(CreateHeadingParagraph("Chapter A", "Heading1")); // 1
            body.Append(CreateSingleCellTable("{{section:4.2}}"));

            var logger = new Mock<ILogger<SectionPlaceholderService>>();
            var service = new SectionPlaceholderService(logger.Object);

            service.ResolveSectionPlaceholders(document);

            var resolvedText = body.Descendants<Text>().Last().Text;
            Assert.Equal("1.4.2", resolvedText);
        }

        [Fact]
        public void ResolveSectionPlaceholders_AnchoredRowsStartAfterExistingSectionNumber()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            body.Append(CreateHeadingParagraph("Chapter 1", "Heading1")); // 1
            body.Append(CreateHeadingParagraph("Chapter 2", "Heading1")); // 2
            body.Append(CreateHeadingParagraph("Chapter 3", "Heading1")); // 3
            body.Append(CreateHeadingParagraph("System Requirements", "Heading1")); // 4
            body.Append(CreateHeadingParagraph("Critical Items/Key Characteristics", "Heading2")); // 4.1
            body.Append(
                new SdtBlock(
                    new SdtProperties(),
                    new SdtContentBlock(new Paragraph(new Run(new Text("{{section-anchor:requirements-root}}"))))
                )
            );

            body.Append(CreateHeadingParagraph("Qualification Provisions", "Heading1")); // 5
            body.Append(CreateSingleCellTable("{{section:requirements-root:1}}"));
            body.Append(CreateSingleCellTable("{{section:requirements-root:2}}"));

            var logger = new Mock<ILogger<SectionPlaceholderService>>();
            var service = new SectionPlaceholderService(logger.Object);

            service.ResolveSectionPlaceholders(document);

            var resolvedValues = body.Descendants<Table>()
                .Select(t => t.Descendants<Text>().First().Text)
                .ToList();

            Assert.Equal("4.2", resolvedValues[0]);
            Assert.Equal("4.3", resolvedValues[1]);
        }

        [Fact]
        public void ResolveSectionPlaceholders_AnchoredRowsFromDeepAnchor_UseChapterAndSecondLevelOffset()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            body.Append(CreateHeadingParagraph("Chapter 1", "Heading1")); // 1
            body.Append(CreateHeadingParagraph("Chapter 2", "Heading1")); // 2
            body.Append(CreateHeadingParagraph("Chapter 3", "Heading1")); // 3
            body.Append(CreateHeadingParagraph("System Requirements", "Heading1")); // 4
            body.Append(CreateHeadingParagraph("Section", "Heading2")); // 4.1
            body.Append(CreateHeadingParagraph("Subsection", "Heading3")); // 4.1.1
            body.Append(CreateHeadingParagraph("Existing Item", "Heading4")); // 4.1.1.1
            body.Append(
                new SdtBlock(
                    new SdtProperties(),
                    new SdtContentBlock(new Paragraph(new Run(new Text("{{section-anchor:requirements-root}}"))))
                )
            );

            body.Append(CreateHeadingParagraph("Qualification Provisions", "Heading1")); // 5
            body.Append(CreateSingleCellTable("{{section:requirements-root:1}}"));
            body.Append(CreateSingleCellTable("{{section:requirements-root:1.4}}"));

            var logger = new Mock<ILogger<SectionPlaceholderService>>();
            var service = new SectionPlaceholderService(logger.Object);

            service.ResolveSectionPlaceholders(document);

            var resolvedValues = body.Descendants<Table>()
                .Select(t => t.Descendants<Text>().First().Text)
                .ToList();

            Assert.Equal("4.2", resolvedValues[0]);
            Assert.Equal("4.2.4", resolvedValues[1]);
        }

        [Fact]
        public void ResolveSectionPlaceholders_NonSysRsAnchor_UsesDepthContinuation()
        {
            using var stream = new MemoryStream();
            using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            body.Append(CreateHeadingParagraph("A", "Heading1")); // 1
            body.Append(CreateHeadingParagraph("B", "Heading2")); // 1.1
            body.Append(CreateHeadingParagraph("C", "Heading3")); // 1.1.1
            body.Append(
                new SdtBlock(
                    new SdtProperties(),
                    new SdtContentBlock(new Paragraph(new Run(new Text("{{section-anchor:custom-anchor}}"))))
                )
            );
            body.Append(CreateSingleCellTable("{{section:custom-anchor:1}}"));
            body.Append(CreateSingleCellTable("{{section:custom-anchor:1.4}}"));

            var logger = new Mock<ILogger<SectionPlaceholderService>>();
            var service = new SectionPlaceholderService(logger.Object);

            service.ResolveSectionPlaceholders(document);

            var resolvedValues = body.Descendants<Table>()
                .Select(t => t.Descendants<Text>().First().Text)
                .ToList();

            Assert.Equal("1.1.2", resolvedValues[0]);
            Assert.Equal("1.1.2.4", resolvedValues[1]);
        }

        [Fact]
        public void ResolveSectionPlaceholders_SysRsTemplate_ResolvesVcrmRowsAsChapterSections()
        {
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.docx");
            var templatePath = FindSysRsTemplatePath();
            if (!string.IsNullOrEmpty(templatePath))
            {
                File.Copy(templatePath, tempPath, true);
            }
            else
            {
                CreateSysRsLikeTemplate(tempPath);
            }

            try
            {
                using var document = WordprocessingDocument.Open(tempPath, true);
                var body = document.MainDocumentPart?.Document?.Body;
                Assert.NotNull(body);

                var systemRequirementsSdt = FindSdtByAlias(body, "system-requirements");
                Assert.NotNull(systemRequirementsSdt);
                var systemRequirementsContent = systemRequirementsSdt.GetFirstChild<SdtContentBlock>();
                Assert.NotNull(systemRequirementsContent);
                systemRequirementsContent.InsertAt(
                    new Paragraph(new Run(new Text("{{section-anchor:requirements-root}}"))),
                    0
                );

                var vcrmSdt = FindSdtByAlias(body, "vcrm");
                Assert.NotNull(vcrmSdt);
                var vcrmContent = vcrmSdt.GetFirstChild<SdtContentBlock>();
                Assert.NotNull(vcrmContent);
                vcrmContent.RemoveAllChildren();
                vcrmContent.Append(CreateSingleCellTable("{{section:requirements-root:1}}"));
                vcrmContent.Append(CreateSingleCellTable("{{section:requirements-root:2}}"));

                var logger = new Mock<ILogger<SectionPlaceholderService>>();
                var service = new SectionPlaceholderService(logger.Object);
                service.ResolveSectionPlaceholders(document);

                var resolvedValues = vcrmSdt.Descendants<Table>()
                    .Select(t => t.Descendants<Text>().First().Text)
                    .ToList();

                Assert.Equal("4.2", resolvedValues[0]);
                Assert.Equal("4.3", resolvedValues[1]);
                Assert.DoesNotContain("{{section-anchor:requirements-root}}", body.InnerText);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        private static string FindSysRsTemplatePath()
        {
            var current = AppContext.BaseDirectory;
            var dir = new DirectoryInfo(current);
            while (dir != null)
            {
                var candidate = Path.Combine(
                    dir.FullName,
                    "docgen-api-gate",
                    "s3-initializer",
                    "assets",
                    "templates",
                    "shared",
                    "SysRS",
                    "SysRS.dotx"
                );
                if (File.Exists(candidate))
                {
                    return candidate;
                }

                dir = dir.Parent;
            }

            return string.Empty;
        }

        private static void CreateSysRsLikeTemplate(string path)
        {
            using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document, true);
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            body.Append(CreateHeadingParagraph("Scope", "Heading1")); // 1
            body.Append(CreateHeadingParagraph("Referenced Documents", "Heading1")); // 2
            body.Append(CreateHeadingParagraph("Engineering Requirements", "Heading1")); // 3
            body.Append(CreateHeadingParagraph("System Requirements", "Heading1")); // 4
            body.Append(CreateHeadingParagraph("Critical Items/Key Characteristics", "Heading2")); // 4.1
            body.Append(
                new SdtBlock(
                    new SdtProperties(
                        new SdtAlias { Val = "system-requirements" },
                        new Tag { Val = "system-requirements" }
                    ),
                    new SdtContentBlock(new Paragraph(new Run(new Text("Click or tap here to enter text."))))
                )
            );
            body.Append(CreateHeadingParagraph("Qualification Provisions", "Heading1")); // 5
            body.Append(
                new SdtBlock(
                    new SdtProperties(
                        new SdtAlias { Val = "vcrm" },
                        new Tag { Val = "vcrm" }
                    ),
                    new SdtContentBlock(new Paragraph(new Run(new Text("Click or tap here to enter text."))))
                )
            );
            mainPart.Document.Save();
        }

        private static SdtBlock FindSdtByAlias(Body body, string alias)
        {
            return body.Descendants<SdtBlock>()
                .FirstOrDefault(s =>
                    string.Equals(
                        s.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value,
                        alias,
                        StringComparison.OrdinalIgnoreCase
                    )
                );
        }

        private static Paragraph CreateHeadingParagraph(string text, string styleId)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = styleId }
                ),
                new Run(new Text(text))
            );
        }

        private static Table CreateSingleCellTable(string text)
        {
            return new Table(
                new TableRow(
                    new TableCell(
                        new Paragraph(new Run(new Text(text)))
                    )
                )
            );
        }
    }
}
