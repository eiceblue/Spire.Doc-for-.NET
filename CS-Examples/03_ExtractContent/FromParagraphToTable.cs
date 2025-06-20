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
            // Create a new source document and load it from a file
            Document sourceDocument = new Document();
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\IncludingTable.docx");

            // Create a new destination document
            Document destinationDoc = new Document();

            // Add a section to the destination document
            Section destinationSection = destinationDoc.AddSection();

            // Extract content by table from the source document to the destination document
            ExtractByTable(sourceDocument, destinationDoc, 1, 1);

            // Save the destination document to a file named "Output.docx"
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            // Dispose of the source and destination documents
            sourceDocument.Dispose();
            destinationDoc.Dispose();

            // Launch the Word file for viewing
            WordDocViewer("Output.docx");
        }

        // Extracts content by table from the source document to the destination document
        private static void ExtractByTable(Document sourceDocument, Document destinationDocument, int startPara, int tableNo)
        {
            // Get the specified table from the source document
            Table table = sourceDocument.Sections[0].Tables[tableNo - 1] as Table;

            // Get the index of the table in the source document
            int index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(table);

            // Copy each child object from the source document to the destination document
            for (int i = startPara - 1; i <= index; i++)
            {
                DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();
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
