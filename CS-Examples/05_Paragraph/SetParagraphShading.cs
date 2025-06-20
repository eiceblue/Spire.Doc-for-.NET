using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetParagraphShading
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Get the first paragraph of the first section in the document.
			Paragraph paragaph = document.Sections[0].Paragraphs[0];

			// Set the background color of the paragraph to yellow.
			paragaph.Format.BackColor = Color.Yellow;

			// Get the third paragraph of the first section in the document.
			paragaph = document.Sections[0].Paragraphs[2];

			// Find the text "Christmas" within the paragraph, starting from the beginning, case-insensitive.
			TextSelection selection = paragaph.Find("Christmas", true, false);

			// Get the found text range as a single range.
			TextRange range = selection.GetAsOneRange();

			// Set the text background color of the range to yellow.
			range.CharacterFormat.TextBackgroundColor = Color.Yellow;

			// Specify the file name for the resulting document.
			string result = "Result-SetParagraphShading.docx";

			// Save the modified document to a file with the specified file name and format (Docx2013).
			document.SaveToFile(result, FileFormat.Docx2013);

			// Clean up resources used by the document.
			document.Dispose();

            //Launch the MS Word file.
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
