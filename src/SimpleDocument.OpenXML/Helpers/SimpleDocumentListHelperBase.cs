using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    /// <summary>Base class for bullet and numbered list.</summary>
    public abstract class SimpleDocumentListHelperBase : SimpleDocumentHelperBase
    {
        public SimpleDocumentListHelperBase(WordprocessingDocument wpd)
        {
            WordprocessingDocument = wpd;
        }

        /// <summary>Add items by not specifying a run (go with default formattting)</summary>
        /// <param name="sentences">List of strings each of which presents a different entry</param>
        public void AddList(List<string> sentences)
        {
            var runList = ParagraphHelper.RunHelper.ConvertToRunList(sentences);

            AddList(runList);
        }

        /// <summary>Add items by not specifying a paragraph (go with default formattting)</summary>
        /// <param name="runList">List of runs each of which presents a different entry</param>
        public void AddList(List<Run> runList)
        {
            var paragraphs = new List<Paragraph>();
            foreach (Run runItem in runList)
            {
                var paragraphProperties = new ParagraphProperties();

                // Spacing
                paragraphProperties.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" }; // Get rid of space between bullets
                // Indentantion
                paragraphProperties.Indentation = new Indentation() { Left = "720", Hanging = "360" }; // correct indentation 
                // Font
                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                RunFonts runFonts1 = new RunFonts() { Ascii = "Symbol", HighAnsi = "Symbol" };
                paragraphMarkRunProperties1.Append(runFonts1);
                paragraphProperties.ParagraphMarkRunProperties = paragraphMarkRunProperties1;
                
                // Create paragraph 
                var newPara = new Paragraph(paragraphProperties);
                newPara.AppendChild(runItem);
                paragraphs.Add(newPara);
            }

            AddList(paragraphs);
        }

        /// <summary>Add bullets by not specifying a run (not concerned about formatting)</summary>
        /// <param name="paragraphs">List of paragraphs each of which presents a different entry</param>
        public void AddList(List<Paragraph> paragraphs)
        {
            var numberId = CreateNumberingEntries();

            foreach (Paragraph theParagraph in paragraphs)
            {
                // If the paragraph has no ParagraphProperties object, create one.
                if (!theParagraph.Elements<ParagraphProperties>().Any())
                {
                    theParagraph.PrependChild(new ParagraphProperties());
                }

                // numberingProperties, spacingBetweenLines1, indentation, paragraphMarkRunProperties1
            // Get the paragraph properties element of the paragraph.
            ParagraphProperties pPr = theParagraph.Elements<ParagraphProperties>().First();

                var numberingProperties = new NumberingProperties(
                    new NumberingLevelReference() { Val = 0 },
                    new NumberingId() { Val = numberId });

                // Do not append since order matters.
                pPr.NumberingProperties = numberingProperties;
                
                ParagraphHelper.AddToBody(theParagraph);
            }
        }
        

        protected abstract int CreateNumberingEntries();
    }
}