using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    /// <summary>Helper to add bulleted lists.</summary>
    /// <remarks>
    /// Helpful links
    /// -  http://stackoverflow.com/questions/1940911/openxml-2-sdk-word-document-create-bulleted-list-programmatically
    /// -  Screen cast on numbering: http://ericwhite.com/blog/screen-cast-wordprocessingml-numbering/
    /// </remarks>
    public class SimpleDocumentBulletListHelper : SimpleDocumentListHelperBase
    {
        public SimpleDocumentBulletListHelper(WordprocessingDocument wpd) : base(wpd) { }


        protected override int CreateNumberingEntries()
        {
            NumberingDefinitionsPart numberingPart = NumberingDefinitionsPartHelper.GetOrCreate();

            // Add AbstractNum
            // Add AbstractNum
            // Add AbstractNum
            var abstractNumberId = numberingPart.Numbering.Elements<AbstractNum>().Count() + 1;
            var abstractNum1 = new AbstractNum { AbstractNumberId = abstractNumberId };

            // Level 1
            var level1 = new Level { LevelIndex = 0 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText1 = new LevelText { Val = "·" };
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            
            // Add it
            abstractNum1.Append(level1);
            NumberingDefinitionsPartHelper.AddAbstractNum(abstractNum1); // Order matters and this method will take care of that.
            

            // Add NumberingInstance
            // Add NumberingInstance
            // Add NumberingInstance
            int numberId = numberingPart.Numbering.Elements<NumberingInstance>().Count() + 1;
            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = numberId };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = abstractNumberId };
            numberingInstance1.Append(abstractNumId1);
            NumberingDefinitionsPartHelper.AddNumberingInstance(numberingInstance1); // Order matters and this method will take care of that.

            return numberId;
        }
    }
}