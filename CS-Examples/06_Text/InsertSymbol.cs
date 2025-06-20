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

namespace InsertSymbol
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

			// Add a new section to the document.
			Section section = document.AddSection();

			// Add a new paragraph to the section.
			Paragraph paragraph = section.AddParagraph();

			// Use a unicode character (U+00C4) to create the symbol Ä and append it to the paragraph.
			TextRange tr = paragraph.AppendText('\u00C4'.ToString());

			// Set the text color of the symbol Ä to red.
			tr.CharacterFormat.TextColor = Color.Red;

			// Append the symbol Ë to the paragraph using a unicode character (U+00CB).
			paragraph.AppendText('\u00CB'.ToString());

			// Specify the file name for the resulting document.
			string result = "Result-InsertSymbol.docx";

			// Save the document to a file with the specified file name and format (Docx 2013).
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
