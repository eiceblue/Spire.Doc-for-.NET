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
            //Create the first document
            Document sourceDocument = new Document();

            //Load the source document from disk.
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\BetweenParagraphStyle.docx");

            //Create a destination document
            Document destinationDoc = new Document();

            //Add a section
            Section section = destinationDoc.AddSection();

            //Extract content between the first paragraph to the third paragraph
            ExtractBetweenParagraphStyles(sourceDocument, destinationDoc, "1", "2");

            //Save the document.
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }

        private static void ExtractBetweenParagraphStyles(Document sourceDocument, Document destinationDocument, string stylename1, string stylename2)
        {
            int startindex = 0;
            int endindex = 0;
            //travel the sections of source document
            foreach (Section section in sourceDocument.Sections)
            {
                //travel the paragraphs
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    //Judge paragraph style1
                    if (paragraph.StyleName == stylename1)
                    {
                        //Get the paragraph index
                        startindex = section.Body.Paragraphs.IndexOf(paragraph);
                    }
                    //Judge paragraph style2
                    if (paragraph.StyleName == stylename2)
                    {
                        //Get the paragraph index
                        endindex = section.Body.Paragraphs.IndexOf(paragraph);
                    }
                }
                //Extract the content
                for (int i = startindex + 1; i < endindex; i++)
                {
                    //Clone the ChildObjects of source document
                    DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

                    //Add to destination document 
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
