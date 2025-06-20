using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertSectionBreak
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
			Document document = new Document();

			// Load a Word document from a specified file path.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Insert a section break at a specific position in the document.
			// There are five section break options: EvenPage, NewColumn, NewPage, NoBreak, OddPage.
			document.Sections[0].Paragraphs[1].InsertSectionBreak(SectionBreakType.NoBreak);

			// Specify the name and file format for the resulting document after saving.
			string result = "Result-InsertSectionBreak.docx";

			// Save the modified document to a file in the specified format (Docx2013).
			document.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources associated with the document.
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
