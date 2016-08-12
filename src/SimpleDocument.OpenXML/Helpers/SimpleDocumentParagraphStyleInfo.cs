namespace SimpleDocument.OpenXML
{
    /// <summary>Class for holder paragrath style info</summary>
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