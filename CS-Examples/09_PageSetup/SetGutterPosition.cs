using System;
using System.Windows.Forms;
using Spire.Doc;

namespace SetGutterPosition
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
			Document document = new Document();

			// Load a Word document from a specified file path.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			// Get the first section of the document.
			Section section = document.Sections[0];

			// Set the top gutter option to true for the section's page setup.
			section.PageSetup.IsTopGutter = true;

			// Set the width of the gutter in points (100f).
			section.PageSetup.Gutter = 100f;

			// Specify the output file name for the modified document.
			string output = "SetGutterPosition.docx";

			// Save the modified document to the specified output file path in DOCX format.
			document.SaveToFile(output, FileFormat.Docx);
			
			// Dispose the existing document object.
			document.Dispose();

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
