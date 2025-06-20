using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SplitText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Specify the input file path.
			string input = @"..\..\..\..\..\..\Data\Sample.docx";

			// Create a new instance of the Document class.
			Document doc = new Document();

			// Load the document from the specified input file.
			doc.LoadFromFile(input);

			// Add a column to the first section of the document with specified widths.
			doc.Sections[0].AddColumn(100f, 20f);

			// Set the "ColumnsLineBetween" property of the page setup in the first section to true, indicating lines between columns.
			doc.Sections[0].PageSetup.ColumnsLineBetween = true;

			// Specify the output file name.
			string output = "SplitText.docx";

			// Save the modified document to a file with the specified output file name and format (Docx2013).
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Clean up resources used by the document.
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
