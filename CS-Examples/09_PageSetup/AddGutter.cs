using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddGutter
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

			// Load a Word document from a specific file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Set the gutter size of the section to 100f (floating-point value)
			section.PageSetup.Gutter = 100f;

			// Specify the output file name
			string output = "InsertGutter.docx";

			// Save the modified document to a file with the specified format (Docx)
			document.SaveToFile(output, FileFormat.Docx);

			// Dispose the Document object to release resources
			document.Dispose();

            //Launching the file
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
