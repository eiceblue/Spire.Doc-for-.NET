using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CompareDocuments
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {   
            // Create a new Document object for the first document
			Document doc1 = new Document();

			// Load the first document from the specified file path
			doc1.LoadFromFile(@"..\..\..\..\..\..\Data\SupportDocumentCompare1.docx");

			// Create a new Document object for the second document
			Document doc2 = new Document();

			// Load the second document from the specified file path
			doc2.LoadFromFile(@"..\..\..\..\..\..\Data\SupportDocumentCompare2.docx");

			// Compare the contents of the two documents and mark differences using "E-iceblue" as the author name
			doc1.Compare(doc2, "E-iceblue");

			// Specify the output file name for the compared result
			string result = "CompareDocuments_result.docx";

			// Save the compared result to the specified file path in Docx2013 format
			doc1.SaveToFile(result, FileFormat.Docx2013);

			// Dispose of the Document objects to release resources
			doc1.Dispose();
			doc2.Dispose();
			
            //View the document
            FileViewer(result);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
