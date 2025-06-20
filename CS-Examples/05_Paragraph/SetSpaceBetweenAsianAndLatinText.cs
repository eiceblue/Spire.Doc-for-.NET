using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetSpaceBetweenAsianAndLatinText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Create a new instance of the Document class.
			Document document = new Document();

			// Load a Word document from a specified file path.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\SetSpaceBetweenAsianAndLatinText.docx");

			// Get the first paragraph of the first section in the document.
			Paragraph para = document.Sections[0].Paragraphs[0];

			// Set whether to automatically adjust space between Asian text and Latin text.
			para.Format.AutoSpaceDE = false;

			// Set whether to automatically adjust space between Asian text and numbers.
			para.Format.AutoSpaceDN = true;

			// Specify the file name for the resulting document.
			string result = "Result.docx";

			// Save the modified document to a file with the specified file name and format (Docx2013).
			document.SaveToFile(result, FileFormat.Docx2013);

			// Clean up resources used by the document.
			document.Dispose();
			
            //Launching the MS Word file.
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
