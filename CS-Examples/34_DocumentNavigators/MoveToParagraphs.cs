using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace MoveToParagraphs
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

            // Move the navigator's cursor to the first section of the document (section index 0).
            navigator.MoveToSection(0);

            // Move the cursor to the third paragraph (index 2) within the current section, at character offset 0.
            navigator.MoveToParagraph(2, 0);

            // Insert new text at the current cursor position, overwriting any existing content from that point onward.
            navigator.Write("This is new content......");

            // Save the modified document to a new file.
            doc.SaveToFile("MoveToParagraphs.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("MoveToParagraphs.docx");
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
