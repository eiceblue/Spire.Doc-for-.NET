using System;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace LoadTextWithEncoding
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
			string inputFile = "../../../../../../Data/Sample_UTF-7.txt";

			// Create a new instance of the Document class.
			Document document = new Document();

			// Load the text content from the specified input file using UTF-7 encoding.
			document.LoadText(inputFile, Encoding.UTF7);

			// Specify the file name for the resulting document.
			string resultFile = "LoadTextWithEncoding_out.docx";

			// Save the loaded text as a Word document with the specified file name and format (Docx).
			document.SaveToFile(resultFile, FileFormat.Docx);

			// Clean up resources used by the document.
			document.Dispose();
			
            WordDocViewer(resultFile);

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
