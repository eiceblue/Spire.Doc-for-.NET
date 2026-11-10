using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace MoveToHeaderAndFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new empty document instance.
            Document doc = new Document();

            // Create a document navigator to help navigate and modify the document content.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Load an existing Word document from the specified relative file path.
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\MoveToHeaderAndFooter.docx");

            // Move the navigator's cursor to the first section of the document (section index 0).
            navigator.MoveToSection(0);

            // Navigate to the footer of the first page in the current section.
            navigator.MoveToHeaderFooter(HeaderFooterType.FooterFirstPage);

            // Write a new line of text into the first-page footer.
            navigator.Writeln("The footer on the first page.");

            // Navigate to the header of the first page in the current section.
            navigator.MoveToHeaderFooter(HeaderFooterType.HeaderFirstPage);

            // Write a new line of text into the first-page header.
            navigator.Writeln("The header on the first page.");

            // Save the modified document to a new file named "MoveToHeaderAndFooter.docx" in DOCX format.
            doc.SaveToFile("MoveToHeaderAndFooter.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("MoveToHeaderAndFooter.docx");
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
