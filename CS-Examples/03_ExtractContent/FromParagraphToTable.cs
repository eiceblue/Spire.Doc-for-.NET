using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace FromParagraphToTable
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
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\IncludingTable.docx");

            //Create a destination document
            Document destinationDoc = new Document();

            //Add a section
            Section destinationSection = destinationDoc.AddSection();

            //Extract the content from the first paragraph to the first table
            ExtractByTable(sourceDocument, destinationDoc, 1, 1);

            //Save the document.
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }

        private static void ExtractByTable(Document sourceDocument, Document destinationDocument, int startPara, int tableNo)
        {
            //Get the table from the source document
            Table table = sourceDocument.Sections[0].Tables[tableNo - 1] as Table;

            //Get the table index
            int index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(table);
            for (int i = startPara - 1; i <= index; i++)
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
