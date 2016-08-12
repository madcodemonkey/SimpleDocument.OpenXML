# SimpleDocument
A simple .NET library for writing to a Microsoft Word document.

Notes
* Since the writer class is performing operations in memory, I do NOT recommend using this code to produce large Word files.
* This has not been released as a NuGet project because it is a work in progress.
* Requires the use of Microsoft's OpenXML library:  DocumentFormat.OpenXML

# Example 
```c#
// Create the writer
var writer = new SimpleDocumentWriter();

// Add a paragraph and adjust its style
var paragraph = writer.ParagraphHelper.AddToBody("This is a good report!");
writer.ParagraphHelper.ApplyStyle(paragraph, SimpleDocumentParagraphStylesEnum.H1);
writer.ParagraphHelper.ApplyJustitification(paragraph, JustificationValues.Center);

// Add a bullet list
List<string> fruitList = new List<string>() { "Apple", "Banana", "Carrot"};
writer.NumberedListHelper.AddList(fruitList);
writer.ParagraphHelper.AddToBody("This is a spacing paragraph 1.");

// Add a numbered list 1, 2, 3
List<string> animalList = new List<string>() { "Dog", "Cat", "Bear" };
writer.NumberedListHelper.AddList(animalList);
writer.ParagraphHelper.AddToBody("This is a spacing paragraph 2.");

// Add an image
using (FileStream fs = new FileStream("C:\\temp\\picture1.jpg", FileMode.Open))
{
	writer.ImageHelper.AddImage(fs);
}

writer.ParagraphHelper.AddToBody(String.Format("This is a spacing paragraph {0}.", pictureNumber));

// Add a paragraph and adjust its style
var addParagraph = writer.ParagraphHelper.AddToBody("Done.");
writer.ParagraphHelper.ApplyStyle(addParagraph, SimpleDocumentParagraphStylesEnum.H1);

// Save the document to a file
writer.SaveToFile("C:\\temp\\Example.docx");
```
