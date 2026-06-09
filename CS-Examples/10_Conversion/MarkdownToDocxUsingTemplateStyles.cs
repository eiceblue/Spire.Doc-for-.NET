using System;
using System.Windows.Forms;
using Spire.Doc;

namespace MarkdownToDocxUsingTemplateStyles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new Document object by loading the Markdown file from the specified relative path.
            Document doc = new Document(@"..\..\..\..\..\..\Data\sample.md");

            // Copy all styles from the specified Word template into the current document.
            doc.CopyStylesFromTemplate(@"..\..\..\..\..\..\Data\template.docx");

            // Define the output filename for the converted Word document.
            String outputFile = "MarkdownToDocxUsingTemplateStyles.docx";

            // Save the processed document to the specified file in DOCX 2016 format.
            doc.SaveToFile(outputFile, FileFormat.Docx2016);

            // Close the document to release resources.
            doc.Close();

            WordDocViewer(outputFile);
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
