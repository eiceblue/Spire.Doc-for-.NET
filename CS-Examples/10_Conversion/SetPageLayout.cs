using System;
using System.Windows.Forms;
using Spire.Doc;

namespace SetPageLayout
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
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\MultiplePages.docx");

            // Instantiate a ToPdfParameterList object to customize PDF export options.
            ToPdfParameterList toPdfParameter = new ToPdfParameterList();

            // Configure the page layout for the output PDF to display two pages per spread, starting with the left page.
            toPdfParameter.PageLayout = PdfPageLayout.TwoPageLeft;

            // Export and save the Word document as a PDF file, applying the previously defined PDF settings.
            document.SaveToFile("SetPageLayout.pdf", toPdfParameter);

            // Close the document to release any file handles or internal resources associated with it.
            document.Close();

            // Immediately free unmanaged resources used by the document object by calling Dispose (good practice for memory management).
            document.Dispose();

            //view the PDF file.
            WordDocViewer("SetPageLayout.pdf");
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
