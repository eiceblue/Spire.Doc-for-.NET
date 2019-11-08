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
            //Create the first document
            Document sourceDocument = new Document();

            //Load the source document from disk.
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            //Create a destination document
            Document destinationDoc = new Document();

            //Add a section
            Section section = destinationDoc.AddSection();

            //Extract content between the first paragraph to the third paragraph
            ExtractBetweenParagraphs(sourceDocument, destinationDoc, 1, 3);

            //Save the document.
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }
        private static void ExtractBetweenParagraphs(Document sourceDocument, Document destinationDocument, int startPara, int endPara)
        {
            //Extract the content
            for (int i = startPara - 1; i < endPara; i++)
            {
                //Clone the ChildObjects of source document
                DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

                //Add to destination document 
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
