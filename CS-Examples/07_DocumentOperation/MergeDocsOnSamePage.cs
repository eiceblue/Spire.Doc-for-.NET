using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace MergeDocsOnSamePage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document document = new Document();

            //Load the source document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Insert.docx");

            //Clone a destination  document
            Document destinationDocument = new Document();

            //Load the destination document from disk.
            destinationDocument.LoadFromFile(@"..\..\..\..\..\..\..\Data\TableOfContent.docx");

            //Traverse sections
            foreach (Section section in document.Sections)
            {

                //Traverse body ChildObjects
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    //Clone to destination document at the same page
                    destinationDocument.Sections[0].Body.ChildObjects.Add(obj.Clone());
                }
            }
            //Save the document.
            destinationDocument.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
