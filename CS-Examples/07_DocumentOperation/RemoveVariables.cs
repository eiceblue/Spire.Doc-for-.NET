using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveVariables
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
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_6.docx");

			// Remove the variable named "A1" from the document's Variables collection
			document.Variables.Remove("A1");

			// Set the IsUpdateFields property of the document to true, enabling field updates
			document.IsUpdateFields = true;

			// Specify the file name for the saved document
			string result = "Result-RemoveVariables.docx";

			// Save the document to a file in DOCX format (using Word 2013 format)
			document.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources used by the document
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
