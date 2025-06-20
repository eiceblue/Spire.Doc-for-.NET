using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetSnapToGrid
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
			Document doc = new Document();

			// Add a new section to the document.
			Section section = doc.AddSection();

			// Set the grid type of the page setup in the section to "LinesOnly".
			section.PageSetup.GridType = GridPitchType.LinesOnly;

			// Set the number of lines per page in the section to 15.
			section.PageSetup.LinesPerPage = 15;

			// Add a new paragraph to the section.
			Paragraph paragraph = section.AddParagraph();

			// Append text to the paragraph.
			paragraph.AppendText("With Spire.Doc, you can generate, modify, convert, render and print documents without utilizing Microsoft Word®. But you need MS Word viewer to view the resultant document. ");

			// Set the "SnapToGrid" property of the paragraph's format to true.
			paragraph.Format.SnapToGrid = true;

			// Specify the file name for the resulting document.
			string output = "SetSnapToGrid.docx";

			// Save the document to a file with the specified file name and format (Docx2013).
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Clean up resources used by the document.
			doc.Dispose();

            //Launch the file 
            WordDocViewer(output);
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

