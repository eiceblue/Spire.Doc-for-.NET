using System;
using System.Windows.Forms;
using Spire.Doc;

namespace HiddenRow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
            Document doc = new Document();

            // Load the content from the specified Word document file path
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\TableTemplate.docx");

            // Get the first section (index 0) from the document's sections collection
            Section section = doc.Sections[0];

            // Cast the first table found in the section to a Table object
            Table table = (Table)section.Tables[0];

            // Get the first row (index 0) from the table's rows collection
            TableRow row = table.Rows[0];

            // Set the Hidden property to true to hide this row in the document
            row.Hidden = true;

            // Define the file path and name for the output document
            String result = "HiddenRow.docx";

            // Save the modified document to a file in standard Docx format
            doc.SaveToFile(result, FileFormat.Docx);

            // Close the document to release resources
            doc.Close();

            // Dispose of the document object to free up memory
            doc.Dispose();

            //Launching the Word file.
            WordDocViewer(result);
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
