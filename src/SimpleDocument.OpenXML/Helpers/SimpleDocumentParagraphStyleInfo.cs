namespace SimpleDocument.OpenXML
{
    internal class SimpleDocumentParagraphStyleInfo
    {
        public SimpleDocumentParagraphStyleInfo()
        {
        }

        public SimpleDocumentParagraphStyleInfo(string styleId, string styleName)
        {
            StyleId = styleId;
            StyleName = styleName;
        }

        public string StyleId { get; set; }
        public string StyleName { get; set; }
    }
}