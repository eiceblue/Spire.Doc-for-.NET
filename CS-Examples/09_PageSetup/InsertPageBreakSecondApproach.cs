using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertPageBreakSecondApproach
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
			Document document = new Document();

			// Load a Word document from a specific file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Get the first section of the document, third paragraph, and append a page break to it
			document.Sections[0].Paragraphs[3].AppendBreak(BreakType.PageBreak);

			// Specify the file name for the resulting document with the inserted page break
			string result = "Result-InsertWordPageBreak.docx";

			// Save the modified document to the specified file path in the DOCX format (version: Word 2013)
			document.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources used by the document object
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
