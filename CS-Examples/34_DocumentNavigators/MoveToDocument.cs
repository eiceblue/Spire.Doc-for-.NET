using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;

namespace MoveToDocument
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

            // Move the cursor to the very beginning of the document.
            navigator.MoveToDocumentStart();

            // Write a new line of text at the start of the document.
            navigator.Writeln("Insert the content at the beginning of the document.");

            // Write another line of text immediately after the previous one at the start.
            navigator.Writeln("This is new content.");

            // Move the cursor to the very end of the document.
            navigator.MoveToDocumentEnd();

            // Insert a blank line at the end of the document.
            navigator.Writeln();

            // Insert a new line of text at the end of the document.
            navigator.Writeln("Insert the content at the end of the document.");

            // Save the modified document to a new file.
            doc.SaveToFile("MoveToDocument.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("MoveToDocument.docx");
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
