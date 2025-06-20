using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace MergeDocsOnSamePage
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

			// Load the source document from a file using a relative path.
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Insert.docx");

			// Create another instance of the Document class.
			Document destinationDocument = new Document();

			// Load the destination document from a file using a relative path.
			destinationDocument.LoadFromFile(@"..\..\..\..\..\..\..\Data\TableOfContent.docx");

			// Iterate through each section in the source document.
			foreach (Section section in document.Sections)
			{
				// Iterate through each child object in the body of the section.
				foreach (DocumentObject obj in section.Body.ChildObjects)
				{
					// Clone each child object and add it to the body of the first section in the destination document.
					destinationDocument.Sections[0].Body.ChildObjects.Add(obj.Clone());
				}
			}

			// Save the destination document to a file with the specified file name and file format as Docx.
			destinationDocument.SaveToFile("Output.docx", FileFormat.Docx);

			// Dispose of the source document and the destination document to release resources.
			document.Dispose();
			destinationDocument.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
