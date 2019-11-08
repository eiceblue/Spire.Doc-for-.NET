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
            //Open a blank Word document as template.
            Document document = new Document(@"..\..\..\..\..\..\Data\Summary_of_Science.doc");

            document.BuiltinDocumentProperties.Title = "Document Demo Document";
            document.BuiltinDocumentProperties.Subject = "demo";
            document.BuiltinDocumentProperties.Author = "James";
            document.BuiltinDocumentProperties.Company = "e-iceblue";
            document.BuiltinDocumentProperties.Manager = "Jakson";
            document.BuiltinDocumentProperties.Category = "Doc Demos";
            document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";
            document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

            //Save as docx file.
            document.SaveToFile("Sample.docx",FileFormat.Docx);

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
