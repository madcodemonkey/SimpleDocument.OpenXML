using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    /// <summary>Helps with adding paragraphs and a LOT of things are in paragraphs (text, images, bullet list items, etc.)</summary>
    public class SimpleDocumentParagraphHelper
    {
        public SimpleDocumentParagraphHelper(WordprocessingDocument wpd)
        {
            WordprocessingDocument = wpd;
        }

        protected Body TheBody
        {
            get
            {
                if (WordprocessingDocument == null)
                    return null;
                return WordprocessingDocument.MainDocumentPart.Document.Body;
            }
        }

        protected WordprocessingDocument WordprocessingDocument { get; set; }

        private SimpleDocumentRunHelper _simpleDocumentRunHelper;

        public SimpleDocumentRunHelper RunHelper
        {
            get { return _simpleDocumentRunHelper ?? (_simpleDocumentRunHelper = new SimpleDocumentRunHelper(WordprocessingDocument)); }
        }

        public Paragraph AddPageBreak()
        {
            return AddToBody(new List<Run> {  RunHelper.CreateBreak(BreakValues.Page) });
        }

        public Paragraph AddToBody(string sentence)
        {
            List<Run> runList = RunHelper.ConvertToRunList(new List<string> { sentence });
            return AddToBody(runList);
        }

        public Paragraph AddToBody(Run sentence)
        {
            var runList = new List<Run> { sentence };
            return AddToBody(runList);
        }

        public Paragraph AddToBody(List<string> sentences)
        {
            List<Run> runList = RunHelper.ConvertToRunList(sentences);
            return AddToBody(runList);
        }

        public Paragraph AddToBody(List<Run> runList)
        {
            var newParagraph = new Paragraph();
            foreach (Run runItem in runList)
            {
                newParagraph.AppendChild(runItem);
            }

            AddToBody(newParagraph);

            return newParagraph;
        }

        /// <summary>Paragraphs should NOT be the last item in the document.  
        /// For example, SectionProperties, which are used with images, should ALWAYS be below the paragraphs! </summary>
        /// <param name="newParagraph"></param>
        public void AddToBody(Paragraph newParagraph)
        {
            var lastParagraph = TheBody.Elements().Any() ? TheBody.Elements<Paragraph>().Last() : null;
            if (lastParagraph == null)
            {
                TheBody.AppendChild(newParagraph);
            }
            else TheBody.InsertAfter(newParagraph, lastParagraph);
        }

        public void ApplyJustitification(Paragraph p, JustificationValues justification)
        {

            // If the paragraph has no ParagraphProperties object, create one.
            if (!p.Elements<ParagraphProperties>().Any())
            {
                p.PrependChild(new ParagraphProperties());
            }

            // Get the paragraph properties element of the paragraph.
            ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

            pPr.Justification = new Justification() { Val = justification };
        }

        /// <summary>Applies a style to a paragraph</summary>
        /// <remarks>
        /// Code from: https://msdn.microsoft.com/en-us/library/office/cc850838.aspx
        /// </remarks>
        public void ApplyStyle(Paragraph p, SimpleDocumentParagraphStylesEnum style)
        {
            if (style == SimpleDocumentParagraphStylesEnum.None)
                return;

            // If the paragraph has no ParagraphProperties object, create one.
            if (!p.Elements<ParagraphProperties>().Any())
            {
                p.PrependChild(new ParagraphProperties());
            }

            // Get the paragraph properties element of the paragraph.
            ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();

            // Get the Styles part for this document.
            StyleDefinitionsPart part = WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart;

            SimpleDocumentParagraphStyleInfo info = GetIntParagraphStyleInfo(style);

            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = AddStylesPartToPackage();
                AddNewStyle(part, style);
            }
            else
            {

                // If the style is not in the document, add it.
                if (IsStyleIdInDocument(info.StyleId) == false)
                {
                    // No match on styleid, so let's try style name.
                    string styleidFromName = GetStyleIdFromStyleName(info.StyleName);
                    if (styleidFromName == null)
                    {
                        AddNewStyle(part, style);
                    }
                    else
                        info.StyleId = styleidFromName;
                }
            }

            // Set the style of the paragraph.
            pPr.ParagraphStyleId = new ParagraphStyleId { Val = info.StyleId };
        }



        private const string NormalStyleId = "Normal";
        private const string NormalStyleName = "Normal";
        private const string Heading1StyleId = "Heading1";
        private const string Heading1StyleName = "heading 1";
        private const string Heading2StyleId = "Heading2";
        private const string Heading2StyleName = "heading 2";

        /// <summary>Create a new style with the specified styleid and stylename and add it to the specified style definitions part.</summary>
        /// <remarks>
        /// Code from: https://msdn.microsoft.com/en-us/library/office/cc850838.aspx
        /// </remarks>
        private void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, SimpleDocumentParagraphStylesEnum styleEnum)
        {
            // Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            Style style;
            switch (styleEnum)
            {
                case SimpleDocumentParagraphStylesEnum.H1:
                    style = CreateH1Style();
                    break;
                case SimpleDocumentParagraphStylesEnum.H2:
                    style = CreateH2Style();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(styleEnum), styleEnum, null);
            }

            if (style != null)
            {
                styles.Append(style);
            }
        }

        /// <summary>Add a StylesDefinitionsPart to the document.  Returns a reference to it.</summary>
        /// <remarks>
        /// Code from: https://msdn.microsoft.com/en-us/library/office/cc850838.aspx
        /// </remarks>
        private StyleDefinitionsPart AddStylesPartToPackage()
        {
            var part = WordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            var root = new Styles();
            root.Save(part);
            return part;
        }

        private Style CreateH1Style()
        {
            Style theStyle = new Style() { Type = StyleValues.Paragraph, StyleId = Heading1StyleId };
            StyleName styleName2 = new StyleName() { Val = Heading1StyleName };
            BasedOn basedOn1 = new BasedOn() { Val = NormalStyleName };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = NormalStyleName };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1Char" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid1 = new Rsid() { Val = "00445B57" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize2 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(runFonts2);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize2);
            styleRunProperties1.Append(fontSizeComplexScript2);

            theStyle.Append(styleName2);
            theStyle.Append(basedOn1);
            theStyle.Append(nextParagraphStyle1);
            theStyle.Append(linkedStyle1);
            theStyle.Append(uIPriority1);
            theStyle.Append(primaryStyle2);
            theStyle.Append(rsid1);
            theStyle.Append(styleRunProperties1);


            return theStyle;
        }

        private Style CreateH2Style()
        {
            Style theStyle = new Style() { Type = StyleValues.Paragraph, StyleId = Heading2StyleId };
            StyleName styleName1 = new StyleName() { Val = Heading2StyleName };
            BasedOn basedOn1 = new BasedOn() { Val = NormalStyleName };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = NormalStyleName };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2Char" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            Rsid rsid1 = new Rsid() { Val = "00445B57" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize1 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            theStyle.Append(styleName1);
            theStyle.Append(basedOn1);
            theStyle.Append(nextParagraphStyle1);
            theStyle.Append(linkedStyle1);
            theStyle.Append(uIPriority1);
            theStyle.Append(rsid1);
            theStyle.Append(styleRunProperties1);
            return theStyle;
        }


        private SimpleDocumentParagraphStyleInfo GetIntParagraphStyleInfo(SimpleDocumentParagraphStylesEnum style)
        {
            switch (style)
            {
                case SimpleDocumentParagraphStylesEnum.None:
                    return new SimpleDocumentParagraphStyleInfo(NormalStyleId, NormalStyleName);
                case SimpleDocumentParagraphStylesEnum.H1:
                    return new SimpleDocumentParagraphStyleInfo(Heading1StyleId, Heading1StyleName);
                case SimpleDocumentParagraphStylesEnum.H2:
                    return new SimpleDocumentParagraphStyleInfo(Heading2StyleId, Heading2StyleName);
                default:
                    throw new ArgumentOutOfRangeException(nameof(style), style, null);
            }
        }


        /// <summary>Return styleid that matches the styleName, or null when there's no match.</summary>
        /// <remarks>
        /// Code from: https://msdn.microsoft.com/en-us/library/office/cc850838.aspx
        /// </remarks>
        private string GetStyleIdFromStyleName(string styleName)
        {
            StyleDefinitionsPart stylePart = WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>().Where(s => s.Val.Value.Equals(styleName) && (((Style)s.Parent).Type == StyleValues.Paragraph)).Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        /// <summary> Return true if the style id is in the document, false otherwise.</summary>
        /// <remarks>
        /// Code from: https://msdn.microsoft.com/en-us/library/office/cc850838.aspx
        /// </remarks>
        private bool IsStyleIdInDocument(string styleid)
        {
            // Get access to the Styles element for this document.
            Styles s = WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;

            // Check that there are styles and how many.
            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            // Look for a match on styleid.
            Style style = s.Elements<Style>().FirstOrDefault(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph));
            if (style == null)
                return false;

            return true;
        }
    }
}
