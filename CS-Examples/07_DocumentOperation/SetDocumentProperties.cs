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
            //Create a document
            Document document = new Document();

            //Load the document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            //Set the build-in Properties.
            document.BuiltinDocumentProperties.Title = "Document Demo Document";
            document.BuiltinDocumentProperties.Author = "James";
            document.BuiltinDocumentProperties.Company = "e-iceblue";
            document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";
            document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

            //Set the custom properties.
            CustomDocumentProperties custom = document.CustomDocumentProperties;
            custom.Add("e-iceblue", true);
            custom.Add("Authorized By", "John Smith");
            custom.Add("Authorized Date", DateTime.Today);

            //Save the document.
            document.SaveToFile("Output.docx", FileFormat.Docx);

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
