using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace KeepSameFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class and load a document from the specified file path ("Template_N2.docx").
			Document srcDoc = new Document();
			srcDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N2.docx");

			// Create a new instance of the Document class and load another document from the specified file path ("Template_N3.docx").
			Document destDoc = new Document();
			destDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N3.docx");

			// Set the KeepSameFormat property of the source document to true.
			srcDoc.KeepSameFormat = true;

			// Iterate through each section in the source document.
			foreach (Section section in srcDoc.Sections)
			{
				// Clone each section and add it to the destination document.
				destDoc.Sections.Add(section.Clone());
			}

			// Specify the output file name.
			string output = "KeepSameFormating_out.docx";

			// Save the modified destination document to a file with the specified output file name and format (Docx2013).
			destDoc.SaveToFile(output, FileFormat.Docx2013);

			// Clean up resources used by the source document.
			srcDoc.Dispose();

			// Clean up resources used by the destination document.
			destDoc.Dispose();

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
