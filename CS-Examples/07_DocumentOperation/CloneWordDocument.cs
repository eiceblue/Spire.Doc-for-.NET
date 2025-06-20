using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CloneWordDocument
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

			// Load a document from the specified file path ("Template_Docx_1.docx").
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Clone the document and assign it to a new Document object.
			Document newDoc = document.Clone();

			// Specify the output file name.
			string result = "Result-CloneWordDocument.docx";

			// Save the cloned document to a file with the specified output file name and format (Docx2013).
			newDoc.SaveToFile(result, FileFormat.Docx2013);

			// Clean up resources used by the original document.
			document.Dispose();

			// Clean up resources used by the cloned document.
			newDoc.Dispose();

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
