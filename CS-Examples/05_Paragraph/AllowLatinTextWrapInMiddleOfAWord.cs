using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AllowLatinTextWrapInMiddleOfAWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
            Document document = new Document();

            // Load an existing document from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\AllowLatinTextWrapInMiddleOfAWord.docx");

            // Get the first paragraph in the first section of the document
            Paragraph para = document.Sections[0].Paragraphs[0];

            // Allow Latin text to wrap in the middle of a word
            para.Format.WordWrap = true;

            // Specify the filename for the resulting document
            string result = "AllowLatinTextWrapInMiddleOfAWord-Result.docx";

            // Save the modified document to the specified file in the Docx2013 format
            document.SaveToFile(result, FileFormat.Docx2013);

            // Dispose of the document resources
            document.Dispose();
          
            //Launching the Word file.
            WordDocViewer(result);
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
