using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Documents.Comparison;
using Spire.Doc.Interface;

namespace CompareDocumentsIgnoreTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Load the first document from the specified file path
            Document document1 = new Document(@"..\..\..\..\..\..\Data\ComparedDoc1.docx");

            // Load the second document from the specified file path
            Document document2 = new Document(@"..\..\..\..\..\..\Data\ComparedDoc2.docx");

            // Create a new CompareOptions object to specify comparison settings
            CompareOptions compareoptions = new CompareOptions();

            // Set the option to ignore differences in tables during comparison
            compareoptions.IgnoreTable = true;

            // Compare the two documents using the specified options, with "E-iceblue" as the author name for tracked changes
            document1.Compare(document2, "E-iceblue", compareoptions);

            // Save the compared document (with changes tracked) to a new file in DOCX 2019 format
            document1.SaveToFile("CompareDocumentsIgnoreTable.docx", FileFormat.Docx2019);

            // Release resources used by the first document
            document1.Dispose();

            // Release resources used by the second document
            document2.Dispose();

            //Launching the MS Word file.
            WordDocViewer("CompareDocumentsIgnoreTable.docx");
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
