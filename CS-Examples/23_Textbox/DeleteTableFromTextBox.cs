using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DeleteTableFromTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
			// Specify the input file path
			string input = @"..\..\..\..\..\..\Data\TextBoxTable.docx";

			// Create a new Document object
			Document doc = new Document();

			// Load a Word document from the specified input file
			doc.LoadFromFile(input);

			// Access the first text box in the document
			Spire.Doc.Fields.TextBox textbox = doc.TextBoxes[0];

			// Remove the table inside the text box
			textbox.Body.Tables.RemoveAt(0);

			// Specify the output file name
			string output = "DeleteTableFromTextBox.docx";

			// Save the modified document to a new file
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to free up resources
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
