using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DocumentProperty
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object and load the document from the specified file path
            Document document = new Document(@"..\..\..\..\..\..\Data\Summary_of_Science.doc");

            // Set the Title property of the document
			document.BuiltinDocumentProperties.Title = "Document Demo Document";

			// Set the Subject property of the document
			document.BuiltinDocumentProperties.Subject = "demo";

			// Set the Author property of the document
			document.BuiltinDocumentProperties.Author = "James";

			// Set the Company property of the document
			document.BuiltinDocumentProperties.Company = "e-iceblue";

			// Set the Manager property of the document
			document.BuiltinDocumentProperties.Manager = "Jakson";

			// Set the Category property of the document
			document.BuiltinDocumentProperties.Category = "Doc Demos";

			// Set the Keywords property of the document
			document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";

			// Set the Comments property of the document
			document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

			// Save the modified document to the specified file path in Docx format
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose of the Document object to release resources
			document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
