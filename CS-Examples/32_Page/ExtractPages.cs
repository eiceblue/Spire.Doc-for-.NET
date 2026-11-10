using System;
using System.Windows.Forms;
using Spire.Doc;

namespace ExtractPages
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class to represent a Word document in memory.
            Document document = new Document();

            // Load content from an existing .docx file located at the specified relative path into the Document object.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\MultiplePages.docx");

            // Extract pages 0 through 2 from the original document and return them as a new Document object.
            Document extractPage = document.ExtractPages(0, 2);

            // Save the extracted pages to a new file.
            extractPage.SaveToFile("ExtractPages.docx", FileFormat.Docx);

            // Close the document to release any file handles or internal resources associated with it.
            document.Close();

            // Immediately free unmanaged resources used by the document object by calling Dispose (good practice for memory management).
            document.Dispose();

            //view the PDF file.
            WordDocViewer("ExtractPages.docx");
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
