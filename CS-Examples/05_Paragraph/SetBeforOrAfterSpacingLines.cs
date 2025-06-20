using System;
using System.Diagnostics;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetBeforOrAfterSpacingLines
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object
            Document doc = new Document();

            // Load a Word document from a specific file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            // Access the first section of the document
            Section section = doc.Sections[0];

            // Access the first paragraph in the section
            Paragraph paragraph = section.Paragraphs[0];

            // Set the spacing before the paragraph 
            paragraph.Format.BeforeSpacingLines = 5f;

            // Set the spacing after the paragraph
            paragraph.Format.AfterSpacingLines = 15f;

            // Save the modified document to a new file
            doc.SaveToFile("setBeforOrAfterSpacingLines.docx");

            // Dispose of the Document object to release resources
            doc.Dispose();

            WordDocViewer("setBeforOrAfterSpacingLines.docx");
        }

        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch(Exception e) {
                Debug.Write(e.StackTrace);
            }
        }

    }
}
