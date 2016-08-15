using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    public class SimpleDocumentRunHelper 
    {
        public SimpleDocumentRunHelper(WordprocessingDocument wpd)
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


        public Run CreateText(string someText)
        {
            var newRun = new Run();
            newRun.AppendChild(new Text(someText));
            return newRun;
        }

        public Run CreateBreak(BreakValues breakType)
        {
            return  new Run(new Break { Type = breakType });
        }

        public List<Run> ConvertToRunList(List<string> sentences)
        {
            var runList = new List<Run>();
            foreach (string item in sentences)
            {
                runList.Add(CreateText(item));
            }

            return runList;
        }

        public void ApplyBold(Run theRun)
        {
            if (theRun.RunProperties == null)
                theRun.RunProperties = new RunProperties();
            theRun.RunProperties.Bold = new Bold();
        }


        public void ApplyUnderline(Run theRun, UnderlineValues underLineType)
        {
            if (theRun.RunProperties == null)
                theRun.RunProperties = new RunProperties();
            theRun.RunProperties.Underline = new Underline() { Val = underLineType };
        }
    }
}
