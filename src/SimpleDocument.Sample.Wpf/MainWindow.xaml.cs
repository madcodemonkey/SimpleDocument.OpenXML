using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Threading;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleDocument.OpenXML;
using Brushes = System.Windows.Media.Brushes;

namespace WordOpenXMLExample1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DoWorkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveToFile("C:\\temp\\Example-file.docx");
                SaveToStream("C:\\temp\\Example-stream.docx");
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }

        private void SaveToFile(string fileToCreate)
        {
            if (File.Exists(fileToCreate))
                File.Delete(fileToCreate);

            var writer = DoWork();

            writer.SaveToFile(fileToCreate);
            LogMessage("File created: " + fileToCreate);
        }

        private void SaveToStream(string fileToCreate)
        {
            if (File.Exists(fileToCreate))
                File.Delete(fileToCreate);

            var writer = DoWork();

            using (FileStream fs = File.Create(fileToCreate))
            using (MemoryStream ms = writer.SaveToStream())
            {
                ms.WriteTo(fs);
            }
            
            LogMessage("File created: " + fileToCreate);
        }

        private SimpleDocumentWriter DoWork()
        {
            var writer = new SimpleDocumentWriter();
            var paragraph1 = writer.ParagraphHelper.AddToBody("This is a good report!");
            writer.ParagraphHelper.ApplyStyle(paragraph1, SimpleDocumentParagraphStylesEnum.H1);
            writer.ParagraphHelper.ApplyJustitification(paragraph1, JustificationValues.Center);

            var paragraph2 = writer.ParagraphHelper.AddToBody("The H2");
            writer.ParagraphHelper.ApplyStyle(paragraph2, SimpleDocumentParagraphStylesEnum.H2);
            writer.ParagraphHelper.ApplyJustitification(paragraph2, JustificationValues.Left);

            List<string> fruitList = new List<string>() {"Apple", "Banana", "Carrot"};
            writer.NumberedListHelper.AddList(fruitList);
            writer.ParagraphHelper.AddToBody("This is a spacing paragraph 1.");

            List<string> animalList = new List<string>() {"Dog", "Cat", "Bear"};
            writer.NumberedListHelper.AddList(animalList);
            writer.ParagraphHelper.AddToBody("This is a spacing paragraph 2.");

            List<string> stuffList = new List<string>() {"Ball", "Wallet", "Phone"};
            writer.BulletHelper.AddList(stuffList);

            AddPicture(writer, @"C:\Temp\picture1.jpg", 1);
            AddPicture(writer, @"C:\Temp\picture2.png", 2);
            AddPicture(writer, @"C:\Temp\picture3.jpg", 3);

            var paragraph3 = writer.ParagraphHelper.AddToBody("Another H1.");
            writer.ParagraphHelper.ApplyStyle(paragraph3, SimpleDocumentParagraphStylesEnum.H1);

            var paragraph4 = writer.ParagraphHelper.AddToBody("The H2");
            writer.ParagraphHelper.ApplyStyle(paragraph4, SimpleDocumentParagraphStylesEnum.H2);
            writer.ParagraphHelper.ApplyJustitification(paragraph4, JustificationValues.Left);

            var paragraph5 = writer.ParagraphHelper.AddToBody("Done");
            writer.ParagraphHelper.ApplyStyle(paragraph5, SimpleDocumentParagraphStylesEnum.H1);
            return writer;
        }

        private void AddPicture(SimpleDocumentWriter writer, string fileNameAndPath, int pictureNumber)
        {
            if (File.Exists(fileNameAndPath))
            {
                using (FileStream fs = new FileStream(fileNameAndPath, FileMode.Open))
                {
                    writer.ImageHelper.AddImage(fs);
                }

                writer.ParagraphHelper.AddToBody(String.Format("This is a spacing paragraph {0}.", pictureNumber));
            }
            else
            {
                LogMessage("Picture not found so it was not added: " + fileNameAndPath);
            }

        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            ClearLog();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveLog();
        }

        #region Logging
        private delegate void NoArgsDelegate();
        private void ClearLog()
        {
            if (Dispatcher.CheckAccess())
            {
                RtbLog.Document.Blocks.Clear();
            }
            else this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new NoArgsDelegate(ClearLog));
        }

        /// <summary>Threadsafe logging method.</summary>
        private void LogMessage(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                var p = new System.Windows.Documents.Paragraph(new System.Windows.Documents.Run(message));
                p.Foreground = Brushes.Black;
                RtbLog.Document.Blocks.Add(p);
            }
            else this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action<string>(LogMessage), message);
        }

        private void LogError(Exception ex)
        {
            if (Dispatcher.CheckAccess())
            {
                // We are back on the UI thread here so calling LogMessage will not cause a BeginInvoke for all these LogMessage calls:
                LogMessage(ex.Message);
                LogMessage(ex.StackTrace);
                if (ex.InnerException != null)
                {
                    LogMessage(ex.InnerException.Message);
                    LogMessage(ex.InnerException.StackTrace);
                }
            }
            else this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action<Exception>(LogError), ex);
        }

        private void SaveLog()
        {
            var dialog = new Microsoft.Win32.SaveFileDialog();
            if (dialog.ShowDialog() != true)
                return;

            using (var fs = new FileStream(dialog.FileName, FileMode.Create))
            {
                var myTextRange = new TextRange(RtbLog.Document.ContentStart, RtbLog.Document.ContentEnd);
                myTextRange.Save(fs, DataFormats.Text);
            }
        }
        #endregion
    }
}
