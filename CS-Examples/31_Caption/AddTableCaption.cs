using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddTableCaption
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
            // Create a new instance of Document
            Document document = new Document();

            // Load the Word document from the specified file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\TableTemplate.docx");

            // Get the body of the first section in the document
            Body body = document.Sections[0].Body;

            // Get the first table in the body
            Table table = body.Tables[0] as Table;

            // Add a caption to the table with the "Table" label, numbering format as "Number", and position below the table
            table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem);

            // Enable field updating in the document
            document.IsUpdateFields = true;

            // Specify the output file name and format (Docx)
            string output = "AddTableCaption_result.docx";
            document.SaveToFile(output, FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launching the file
            WordDocViewer(output);

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
