using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    /// <summary>Helps with NumberingDefinitionsPart to add NumberingInstance add AbstractNum since there order matters this takes care of that problem.</summary>
    public class SimpleDocumentNumberingDefinitionsPartHelper
    {
        public SimpleDocumentNumberingDefinitionsPartHelper(WordprocessingDocument wpd)
        {
            WordprocessingDocument = wpd;
        }
        public Body TheBody
        {
            get
            {
                if (WordprocessingDocument == null)
                    return null;
                return WordprocessingDocument.MainDocumentPart.Document.Body;
            }
        }
        public WordprocessingDocument WordprocessingDocument { get; set; }

        /// <summary>Insert an NumberingInstance into the numbering part numbering list.  The order seems to matter 
        /// or it will not pass the Open XML SDK Productity Tools validation method.  AbstractNum comes first and
        /// then NumberingInstance and we want to insert this AFTER the last NumberingInstance and AFTER all the 
        /// AbstractNum entries or we will get a validation error.</summary>
        /// <param name="newNumberingInstance">Item to add</param>
        public void AddNumberingInstance(NumberingInstance newNumberingInstance)
        {
            NumberingDefinitionsPart numberingPart = GetOrCreate();

            if (numberingPart.Numbering.Elements<NumberingInstance>().Any())
            {
                var lastNumberingInstance = numberingPart.Numbering.Elements<NumberingInstance>().Last();
                numberingPart.Numbering.InsertAfter(newNumberingInstance, lastNumberingInstance);
            }
            else
            {
                numberingPart.Numbering.Append(newNumberingInstance);
            }
        }


        /// <summary>Insert an AbstractNum into the numbering part numbering list.  The order seems to matter
        /// or it will not pass the Open XML SDK Productity Tools validation method.  AbstractNum comes first 
        /// and then NumberingInstance and we want to insert this AFTER the last AbstractNum and BEFORE the 
        /// first NumberingInstance or we will get a validation error.</summary>
        /// <param name="newAbstractNum">Item to add</param>
        public void AddAbstractNum(AbstractNum newAbstractNum)
        {
            NumberingDefinitionsPart numberingPart = GetOrCreate();

            if (numberingPart.Numbering.Elements<AbstractNum>().Any())
            {
                AbstractNum lastAbstractNum = numberingPart.Numbering.Elements<AbstractNum>().Last();
                numberingPart.Numbering.InsertAfter(newAbstractNum, lastAbstractNum);
            }
            else
            {
                numberingPart.Numbering.Append(newAbstractNum);
            }
        }

        public NumberingDefinitionsPart GetOrCreate()
        {
            // Introduce bulleted numbering in case it will be needed at some point
            NumberingDefinitionsPart numberingPart = WordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = WordprocessingDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");
                Numbering element = new Numbering();
                element.Save(numberingPart);
            }

            return numberingPart;
        }

    }
}
