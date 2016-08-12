using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    public class SimpleDocumentImageHelper : SimpleDocumentHelperBase
    {
        public SimpleDocumentImageHelper(WordprocessingDocument wpd)
        {
            WordprocessingDocument = wpd;
        }

        /// <summary>Add an image to the body of the document.  This method is used when data is retrieved from a database as a byte array.</summary>
        /// <param name="imageBytes">Image as a byte array</param>
        public Paragraph AddImage(byte[] imageBytes)
        {
            using (var ms = new MemoryStream(imageBytes))
            {
                return AddImage(ms);
            }
        }

        /// <summary>Add an image to the body of the document.</summary>
        /// <remarks>Based on this code:  https://msdn.microsoft.com/en-us/library/bb497430.aspx</remarks>
        /// <param name="fileData">Stream containing the image.</param>
        public Paragraph AddImage(Stream fileData)
        {
            // Document size is DXA or 1/20 if a point.
            // 1 DXA = 1/20 of a point
            // 20 DXA = 1 pt
            // 12700 EMUs = 1pt
            // 914400 EMUs = 1 inch

            var sectionProperties = TheBody.GetFirstChild<SectionProperties>();
            if (sectionProperties == null)
            {
                sectionProperties = new SectionProperties() { RsidR = "00616CEE" };
                PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
                PageMargin pageMargin1 = new PageMargin()
                {
                    Top = 1440,
                    Right = (UInt32Value)1440U,
                    Bottom = 1440,
                    Left = (UInt32Value)1440U,
                    Header = (UInt32Value)720U,
                    Footer = (UInt32Value)720U,
                    Gutter = (UInt32Value)0U
                };
                Columns columns1 = new Columns() { Space = "720" };
                DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

                sectionProperties.Append(pageSize1);
                sectionProperties.Append(pageMargin1);
                sectionProperties.Append(columns1);
                sectionProperties.Append(docGrid1);
                TheBody.Append(sectionProperties);
            }


            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            //var maxPageWidth = 7315200;// pageSize.Width * 6350; // OR (pageSize.Width/20) * 12700;

            // Get margins: http://stackoverflow.com/questions/21490343/detect-printable-area-width-in-openxml-wordprocessing

            // this contains information about surrounding margins
            var pageMargin = sectionProperties.GetFirstChild<PageMargin>();
            var maxWithInDxas = pageSize.Width - pageMargin.Right.Value - pageMargin.Left.Value;
            var maxPageWidthInEmus = maxWithInDxas * 635;
            //var maxPageWidth =(long)(8.5 * 914400) - ((pageMargin.Right.Value - pageMargin.Left.Value) * 6350); // OR (pageSize.Width/20) * 12700;
            // So if page width is 12240 


            var img = new Bitmap(fileData);
            var widthPx = img.Width;
            var heightPx = img.Height;
            var horzRezDpi = img.HorizontalResolution;
            var vertRezDpi = img.VerticalResolution;
            const int emusPerInch = 914400;
            var widthEmus = (long)((widthPx / horzRezDpi) * emusPerInch);
            var heightEmus = (long)((heightPx / vertRezDpi) * emusPerInch);


            if (widthEmus > maxPageWidthInEmus)
            {
                var ratio = (heightEmus * 1.0m) / widthEmus;
                widthEmus = maxPageWidthInEmus;
                heightEmus = (long)(widthEmus * ratio);
            }

            fileData.Position = 0;
            ImagePart imagePart = WordprocessingDocument.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
            imagePart.FeedData(fileData);


            return AddImageToBody(WordprocessingDocument.MainDocumentPart.GetIdOfPart(imagePart), heightEmus, widthEmus);
        }

        private Paragraph AddImageToBody(string relationshipId, long heightInEmus, long widthInEmus)
        {
            uint nextId = GetMaxDocPropertyId() + 1;

            // Define the reference of the image.
            var element =
                new Drawing(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = widthInEmus, Cy = heightInEmus },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                        {
                            Id = (UInt32Value)nextId,
                            Name = string.Format("Picture {0}", nextId)
                        },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                            new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = "New Bitmap Image.jpg"
                                        },
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new DocumentFormat.OpenXml.Drawing.Blip(
                                            new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                                                new DocumentFormat.OpenXml.Drawing.BlipExtension()
                                                {
                                                    Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                })
                                            )
                                        {
                                            Embed = relationshipId,
                                            CompressionState =
                                                DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                        },
                                        new DocumentFormat.OpenXml.Drawing.Stretch(
                                            new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                        new DocumentFormat.OpenXml.Drawing.Transform2D(
                                            new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                            new DocumentFormat.OpenXml.Drawing.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
                                        new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                            new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                            )
                                        { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                                )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U
                    });

            // Append the reference to body, the element should be in a Run.
            var newParagraph = new Paragraph(new Run(element));
            ParagraphHelper.AddToBody(newParagraph);

            return newParagraph;
        }

    }
}
