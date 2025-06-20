using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace LinkHeadersFooters
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
			Document srcDoc = new Document();

			// Load the source document from a file using a relative path.
			srcDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N1.docx");

			// Create another instance of the Document class.
			Document dstDoc = new Document();

			// Load the destination document from a file using a relative path.
			dstDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N2.docx");

			// Link the header of the first section in the source document to the previous section's header.
			srcDoc.Sections[0].HeadersFooters.Header.LinkToPrevious = true;

			// Link the footer of the first section in the source document to the previous section's footer.
			srcDoc.Sections[0].HeadersFooters.Footer.LinkToPrevious = true;

			// Iterate through each section in the source document.
			foreach (Section section in srcDoc.Sections)
			{
				// Clone each section and add it to the destination document.
				dstDoc.Sections.Add(section.Clone());
			}

			// Specify the output file name.
			string output = "LinkHeadersFooters_out.docx";

			// Save the destination document to a file with the specified file format (Docx2013).
			dstDoc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose of the source and destination documents to free up resources.
			srcDoc.Dispose();
			dstDoc.Dispose();

            //Launching the document
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
