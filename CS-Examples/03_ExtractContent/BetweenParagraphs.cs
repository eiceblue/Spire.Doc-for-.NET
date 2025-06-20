using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace BetweenParagraphs
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
             // Create a new instance of the Document class.
            Document sourceDocument = new Document();

            // Load a document from the given file path. The path is relative to the current executable directory.
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            // Create a new instance of the Document class.
            Document destinationDoc = new Document();

            // Add a new section to the destination document.
            Section section = destinationDoc.AddSection();

            // Call the ExtractBetweenParagraphs function to extract text between specified paragraphs from the source document and add it to the destination document.
            ExtractBetweenParagraphs(sourceDocument, destinationDoc, 1, 3);

            // Save the modified document to the given file name with the .docx file format.
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            // Dispose the sourceDocument and destinationDoc to release the associated resources.
            sourceDocument.Dispose();
            destinationDoc.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }
		
        // This function clones the text between the start and end paragraphs from the source document and adds it to the destination document.
        private static void ExtractBetweenParagraphs(Document sourceDocument, Document destinationDocument, int startPara, int endPara)
        {
            for (int i = startPara - 1; i < endPara; i++)
            {
                // Clone the paragraph object from the source document.
                DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

                // Add the cloned paragraph object to the destination document.
                destinationDocument.Sections[0].Body.ChildObjects.Add(doobj);
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
