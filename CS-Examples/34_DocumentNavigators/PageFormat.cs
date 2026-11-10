using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace PageFormat
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
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            // Move the navigator's cursor to the first section (section index 0) of the document.
            navigator.MoveToSection(0);

            // Set the page margins for the current section.
            navigator.PageSetup.Margins = new MarginsF(100, 80, 100, 80);

            // Set the page size of the current section to Letter.
            navigator.PageSetup.PageSize = PageSize.Letter;

            // Save the modified document to a new file.
            doc.SaveToFile("PageFormat.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("PageFormat.docx");
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
