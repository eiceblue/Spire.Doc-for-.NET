using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Interface;

namespace ReplaceWithDocument
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Load a template document 
            Document doc = new Document(@"..\..\..\..\..\..\Data\Text2.docx");

            //Load another document to replace text
            IDocument replaceDoc = new Document(@"..\..\..\..\..\..\Data\Text1.docx");

            //Replace specified text with the other document
            doc.Replace("Document1", replaceDoc, false, true);

            //Save and launch document
            string output = "ReplaceWithDocument.docx";
            doc.SaveToFile(output, FileFormat.Docx);
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
