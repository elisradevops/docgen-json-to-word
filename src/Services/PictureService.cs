using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using JsonToWord.Services.Interfaces;
using SixLabors.ImageSharp;
using System;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace JsonToWord.Services
{
    public class PictureService : IPictureService
    {
        private readonly IContentControlService _contentControlService;
        private readonly IParagraphService _paragraphService;
        private uint _currentId = 1;

        public PictureService(IContentControlService contentControlService, IParagraphService paragraphService)
        {
            _contentControlService = contentControlService;
            _paragraphService = paragraphService;
        }
        public void Insert(WordprocessingDocument document, string contentControlTitle, WordAttachment wordAttachment)
        {
            _currentId = GetMaxImageId(document) + 1;
            var drawing = CreateDrawing(document.MainDocumentPart, wordAttachment.Path);

            var run = new Run();
            run.AppendChild(drawing);

            var paragraph = new Paragraph();
            paragraph.AppendChild(run);

            // Create and add the caption below the image
            var captionParagraph = _paragraphService.CreateCaption(wordAttachment.Name);

            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(paragraph);
            sdtContentBlock.AppendChild(captionParagraph);

            sdtBlock.AppendChild(sdtContentBlock);
        }

        public Drawing CreateDrawing(MainDocumentPart mainDocumentPart, string filePath, bool isFlattened = false)
        {
            var imagePartId = AddImagePart(mainDocumentPart, filePath);
            var drawingExtend = GetDrawingExtend(filePath, isFlattened);
            uint uniqueId = _currentId++;

            var inline = new DW.Inline(
                new DW.Extent { Cx = drawingExtend.Width, Cy = drawingExtend.Height },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = uniqueId, Name = $"Picture {uniqueId}" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = "New Bitmap Image.jpg" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip(
                                    new A.BlipExtensionList(
                                        new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }
                                    )
                                )
                                { Embed = imagePartId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = drawingExtend.Width, Cy = drawingExtend.Height }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            };

            return new Drawing(inline);
        }

        private uint GetMaxImageId(WordprocessingDocument document)
        {
            var ids = document.MainDocumentPart.Document.Body
                .Descendants<DW.DocProperties>()
                .Select(dp => dp.Id.Value)
                .Union(document.MainDocumentPart.Document.Body
                .Descendants<PIC.NonVisualDrawingProperties>()
                .Select(nvdp => nvdp.Id.Value));

            return ids.Any() ? ids.Max() : 1;
        }

        private static DrawingExtent GetDrawingExtend(string localPath, bool isFlattened = false)
        {
            int width, height;

            using (var bmp = Image.Load(localPath))
            {
                width = bmp.Width;
                height = bmp.Height;
            }

            const int maxWidth = 5715000;
            const int scaleFactor = 9525;

            width = Math.Min((int)Math.Round((decimal)width * scaleFactor), maxWidth);
            height = (int)Math.Round((decimal)height * scaleFactor);

            return isFlattened
                ? new DrawingExtent(height / 2, width / 2)
                : new DrawingExtent(height, width);
        }


        private string AddImagePart(MainDocumentPart mainDocumentPart, string imagePath)
        {
            var imagePart = mainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            return mainDocumentPart.GetIdOfPart(imagePart);
        }

    }
}