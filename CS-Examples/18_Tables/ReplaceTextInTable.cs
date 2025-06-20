using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace ReplaceTextInTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
   
			// Create a new document object
			Document doc = new Document();

			// Load a document from a file, specified by the file path
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextInTable.docx");

			// Get the first section of the document
			Section section = doc.Sections[0];

			// Get the first table in the section
			Table table = section.Tables[0] as Table;

			// Create a regular expression pattern for matching text within curly braces
			System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"{[^\}]+\}");

			// Replace text in the table that matches the regular expression pattern with "E-iceblue"
			table.Replace(regex, "E-iceblue");

			// Replace the text "Beijing" with "Component" in the table, case-insensitive and match whole words only
			table.Replace("Beijing", "Component", false, true);

			// Specify the output file name
			string output = "ReplaceTextInTable_out.docx";

			// Save the modified document to a file, using Docx2013 format
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose of the document object
			doc.Dispose();

            //Launch the file
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
