using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetDocumentProperties
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
			Document document = new Document();

			// Load the document from the specified file path using a relative path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			// Set the values of various built-in document properties
			document.BuiltinDocumentProperties.Title = "Document Demo Document";
			document.BuiltinDocumentProperties.Author = "James";
			document.BuiltinDocumentProperties.Company = "e-iceblue";
			document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";
			document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

			// Get the collection of custom document properties
			CustomDocumentProperties custom = document.CustomDocumentProperties;

			// Add custom document properties to the collection
			custom.Add("e-iceblue", true);
			custom.Add("Authorized By", "John Smith");
			custom.Add("Authorized Date", DateTime.Today);

			// Save the modified document to a file with the specified output file name and file format (Docx)
			document.SaveToFile("Output.docx", FileFormat.Docx);

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
