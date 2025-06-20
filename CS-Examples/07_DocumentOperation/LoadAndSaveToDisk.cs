using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
namespace LoadAndSaveToDisk
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define the input file path using a relative path
			String input = @"..\..\..\..\..\..\Data\Template.docx";

			// Create a new instance of the Document class
			Document doc = new Document();

			// Load the document from the specified input file
			doc.LoadFromFile(input);

			// Specify the output file name
			String result = "LoadAndSaveToDisk_out.docx";

			// Save the document to a file with the specified output file name and file format (Docx)
			doc.SaveToFile(result, FileFormat.Docx);

			// Dispose of the document object to free up resources
			doc.Dispose();
			
            WordViewer(result);
        }
        private void WordViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
