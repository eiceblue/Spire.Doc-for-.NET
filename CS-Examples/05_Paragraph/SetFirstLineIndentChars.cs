using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetFirstLineIndentChars
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

			// Load a Word document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Create a Paragraph object using the loaded document
			Paragraph para = new Paragraph(document);

			// Append text to the paragraph and customize its formatting
			TextRange textRange1 = para.AppendText("This is an inserted paragraph.");
			textRange1.CharacterFormat.TextColor = Color.Blue;
			textRange1.CharacterFormat.FontSize = 15;

			// Set the first line indent of the paragraph to 2 characters
			para.Format.FirstLineIndentChars = 2;

			// Alternatively, set the hanging indent as 2 characters
			// para.Format.FirstLineIndentChars = -2;

			// Reset the first line indent to 0 characters
			para.Format.SetFirstLineIndentChars(0);

			// Insert the paragraph at index 1 in the first section of the document
			document.Sections[0].Paragraphs.Insert(1, para);

			// Save the modified document to a new file named "Result-SetFirstLineIndentChars.docx"
			string result = "Result-SetFirstLineIndentChars.docx";
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose the Document object to release resources
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
