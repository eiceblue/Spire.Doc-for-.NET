using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace BetweenParagraphStyles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new source document
            Document sourceDocument = new Document();

            // Load the source document from a file
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\BetweenParagraphStyle.docx");

            // Create a new destination document
            Document destinationDoc = new Document();

            // Add a section to the destination document
            Section section = destinationDoc.AddSection();

            // Extract paragraphs between specified styles from the source document and copy them to the destination document
            ExtractBetweenParagraphStyles(sourceDocument, destinationDoc, "1", "2");

            // Save the destination document to a file named "Output.docx"
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            // Dispose the sourceDocument and destinationDoc to release the associated resources.
            sourceDocument.Dispose();
            destinationDoc.Dispose();

            // Open the Word file
            WordDocViewer("Output.docx");
        }

        // Method to extract paragraphs between two paragraph styles
        private static void ExtractBetweenParagraphStyles(Document sourceDocument, Document destinationDocument, string stylename1, string stylename2)
        {
            int startindex = 0;
            int endindex = 0;

            // Iterate through sections in the source document
            foreach (Section section in sourceDocument.Sections)
            {
                // Iterate through paragraphs in the section
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    // Find the starting paragraph style
                    if (paragraph.StyleName == stylename1)
                    {
                        startindex = section.Body.Paragraphs.IndexOf(paragraph);
                    }

                    // Find the ending paragraph style
                    if (paragraph.StyleName == stylename2)
                    {
                        endindex = section.Body.Paragraphs.IndexOf(paragraph);
                    }
                }

                // Copy paragraphs between the starting and ending indexes
                for (int i = startindex + 1; i < endindex; i++)
                {
                    // Clone the document object
                    DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

                    // Add the cloned object to the destination document
                    destinationDocument.Sections[0].Body.ChildObjects.Add(doobj);
                }
            }
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
