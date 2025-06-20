using System;
using System.Windows.Forms;
using Spire.Doc;

namespace MarkDownToWordOrPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define the input file path relative to the current directory.
            string input = @"..\..\..\..\..\..\Data\MarkDownFile.md";

            // Create a new Document object
            Document doc = new Document();

            //Load .md file
            doc.LoadFromFile(input);

            //Save to .md file
            //doc.SaveToFile("output.md", Spire.Doc.FileFormat.Markdown);
            //Save to .docx file
            //doc.SaveToFile("output.docx", Spire.Doc.FileFormat.Docx);
            //Save to .doc file
            //doc.SaveToFile("output.doc", Spire.Doc.FileFormat.Doc);
            //Save to .pdf file
            doc.SaveToFile("output.pdf",FileFormat.PDF);

            // Dispose of the Document object
            doc.Close();

       
            Viewer("output.pdf");
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
