using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CopyContentToAnotherDoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class and load a document from the specified file path ("Template_Docx_1.docx").
			Document sourceDoc = new Document();
			sourceDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_1.docx");

			// Create a new instance of the Document class and load another document from the specified file path ("Target.docx").
			Document destinationDoc = new Document();
			destinationDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Target.docx");

			// Iterate through each section in the source document.
			foreach (Section sec in sourceDoc.Sections)
			{
				// Iterate through each child object in the body of the section.
				foreach (DocumentObject obj in sec.Body.ChildObjects)
				{
					// Clone the child object and add it to the body of the first section in the destination document.
					destinationDoc.Sections[0].Body.ChildObjects.Add(obj.Clone());
				}
			}

			// Specify the output file name.
			string result = "Result-CopyContentToAnotherWord.docx";

			// Save the modified destination document to a file with the specified output file name and format (Docx2013).
			destinationDoc.SaveToFile(result, FileFormat.Docx2013);

			// Clean up resources used by the source document.
			sourceDoc.Dispose();

			// Clean up resources used by the destination document.
			destinationDoc.Dispose();

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
