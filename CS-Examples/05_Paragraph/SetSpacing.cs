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

namespace SetSpacing
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

			// Create a new paragraph object and associate it with the document.
			Paragraph para = new Paragraph(document);

			// Append text to the paragraph and apply formatting properties.
			TextRange textRange1 = para.AppendText("This is an inserted paragraph.");
			textRange1.CharacterFormat.TextColor = Color.Blue;
			textRange1.CharacterFormat.FontSize = 15;

			// Disable automatic spacing before the paragraph.
			para.Format.BeforeAutoSpacing = false;
			// Set the amount of spacing before the paragraph to 10 points.
			para.Format.BeforeSpacing = 10;
			// Disable automatic spacing after the paragraph.
			para.Format.AfterAutoSpacing = false;
			// Set the amount of spacing after the paragraph to 10 points.
			para.Format.AfterSpacing = 10;

			// Insert the newly created paragraph at index 1 within the paragraphs collection of the first section in the document.
			document.Sections[0].Paragraphs.Insert(1, para);

			// Specify the file name for the resulting document.
			string result = "Result-SetTheSpacing.docx";

			// Save the modified document to a file with the specified file name and format (Docx2013).
			document.SaveToFile(result, FileFormat.Docx2013);

			// Clean up resources used by the document.
			document.Dispose();

            //Launch the file.
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
