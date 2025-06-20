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

           // Create a new Word document object with a relative path
            Document doc = new Document(@"..\..\..\..\..\..\Data\Text2.docx");

            // Create an object of another Word document to be used for replacement
            IDocument replaceDoc = new Document(@"..\..\..\..\..\..\Data\Text1.docx");

            // Search for the string "Document1" in the doc document for the first occurrence,
            // replace it with the content of the replaceDoc document, case-sensitive search,
            // but case-insensitive replacement
            doc.Replace("Document1", replaceDoc, false, true);

            // Define the output file name, which will be saved in the root directory of the project with a .docx extension
            string output = "ReplaceWithDocument.docx";

            // Save the modified document to a specified path with a .docx format
            doc.SaveToFile(output, FileFormat.Docx);

            // Dispose of the document object to release resources and prevent memory leaks
            doc.Dispose();
			
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
