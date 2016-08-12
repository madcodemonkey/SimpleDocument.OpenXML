using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    /// <summary>Abstract base class for all helpers.</summary>
    public abstract class SimpleDocumentHelperBase
    {
        private SimpleDocumentParagraphHelper _paragraphHelper;
        private SimpleDocumentNumberingDefinitionsPartHelper _numberingDefinitionsPartHelper;

        public SimpleDocumentNumberingDefinitionsPartHelper NumberingDefinitionsPartHelper
        {
            get { return _numberingDefinitionsPartHelper ?? (_numberingDefinitionsPartHelper = new SimpleDocumentNumberingDefinitionsPartHelper(WordprocessingDocument)); }
        }

        public SimpleDocumentParagraphHelper ParagraphHelper
        {
            get { return _paragraphHelper ?? (_paragraphHelper = new SimpleDocumentParagraphHelper(WordprocessingDocument)); }
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

        public WordprocessingDocument WordprocessingDocument { get; set; }


        /// <summary>Finds the largest Id number for a DocProperties object.  DocProperties could be buried 
        /// in a number of paragraphs.  This scans the entire document looking for them.</summary>
        /// <returns>Max Id</returns>
        protected uint GetMaxDocPropertyId()
        {
            return WordprocessingDocument
                .MainDocumentPart
                .RootElement
                .Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
                .Max(x => (uint?)x.Id) ?? 0;
        }
    }
}