using System;
using System.Windows.Forms;
using Spire.Doc;

namespace Merge
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

				// Load the source document from a file using a relative path and specify the file format as Doc.
				document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Summary_of_Science.doc", FileFormat.Doc);

				// Create another instance of the Document class.
				Document documentMerge = new Document();

				// Load the document to merge from a file using a relative path and specify the file format as Docx.
				documentMerge.LoadFromFile(@"..\..\..\..\..\..\..\Data\Bookmark.docx", FileFormat.Docx);

				// Iterate through each section in the document to merge.
				foreach (Section sec in documentMerge.Sections)
				{
					// Clone each section from the document to merge and add it to the main document.
					document.Sections.Add(sec.Clone());
				}

				// Save the merged document to a file with the specified file name and file format as Docx.
				document.SaveToFile("Sample.docx", FileFormat.Docx);

				// Dispose of the main document and the document to merge to release resources.
				document.Dispose();
				documentMerge.Dispose();

                //Launching the MS Word file.
                WordDocViewer("Sample.docx");


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
