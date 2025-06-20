using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddLineNumbers
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

			// Set the start value for line numbering in the first section of the document
			document.Sections[0].PageSetup.LineNumberingStartValue = 1;

			// Set the interval between line numbers in the first section of the document
			document.Sections[0].PageSetup.LineNumberingStep = 6;

			// Set the distance between line numbers and the main text in the first section of the document
			document.Sections[0].PageSetup.LineNumberingDistanceFromText = 40f;

			// Set the line numbering restart mode to continuous in the first section of the document
			document.Sections[0].PageSetup.LineNumberingRestartMode = LineNumberingRestartMode.Continuous;

			// Specify the file name for the resulting document with line numbers
			string result = "Result-AddLineNumbers.docx";

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
