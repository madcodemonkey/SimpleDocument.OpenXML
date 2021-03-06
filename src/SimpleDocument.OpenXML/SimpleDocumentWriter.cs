﻿using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleDocument.OpenXML
{
    public class SimpleDocumentWriter : SimpleDocumentHelperBase, IDisposable
    {
        private MemoryStream _ms;
        private SimpleDocumentBulletListHelper _bulletedList;
        private SimpleDocumentNumberListHelper _numberedList;
        private SimpleDocumentImageHelper _imageHelper;

        public SimpleDocumentWriter()
        {
            Reset();
        }

        public SimpleDocumentBulletListHelper BulletHelper
        {
            get { return _bulletedList ?? (_bulletedList = new SimpleDocumentBulletListHelper(WordprocessingDocument)); }
        }

        public SimpleDocumentNumberListHelper NumberedListHelper
        {
            get { return _numberedList ?? (_numberedList = new SimpleDocumentNumberListHelper(WordprocessingDocument)); }
        }
        public SimpleDocumentImageHelper ImageHelper
        {
            get { return _imageHelper ??(_imageHelper = new SimpleDocumentImageHelper(WordprocessingDocument)); }
        }

        public void Reset()
        {
            CloseAndDisposeOfDocument();

            if (_ms == null)
                _ms = new MemoryStream();
            else _ms.Position = 0;

            // These helpers have a reference to the WordprocessingDocument so null them out!
            _bulletedList = null;
            _numberedList = null;
            _imageHelper = null;

            // Create a new one
            WordprocessingDocument = WordprocessingDocument.Create(_ms, WordprocessingDocumentType.Document);
            var mainDocumentPart = WordprocessingDocument.AddMainDocumentPart();
            mainDocumentPart.Document = new Document(new Body());
        }

        public void Dispose()
        {
            CloseAndDisposeOfDocument();
            if (_ms != null)
            {
                _ms.Dispose();
                _ms = null;
            }
        }

        public MemoryStream SaveToStream()
        {
            CloseAndDisposeOfDocument();

            _ms.Position = 0;
            return _ms;
        }

        public void SaveToFile(string fileName)
        {
            CloseAndDisposeOfDocument();

            if (_ms == null)
                throw new ArgumentException("This object has already been disposed of so you cannot save it!");

            using (var fs = File.Create(fileName))
            {
                _ms.WriteTo(fs);
            }
        }

        private void CloseAndDisposeOfDocument()
        {
            if (WordprocessingDocument != null)
            {
                WordprocessingDocument.Close();
                WordprocessingDocument.Dispose();
                WordprocessingDocument = null;
            }
        }
    }
}
