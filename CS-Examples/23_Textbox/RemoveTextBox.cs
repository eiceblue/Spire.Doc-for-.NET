using System;
using System.Windows.Forms;
using Spire.Doc;

namespace RemoveTextBox
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
			string input = @"..\..\..\..\..\..\Data\TextBoxTemplate.docx";

			// Create a new instance of Document
			Document doc = new Document();

			// Load the document from the specified input file
			doc.LoadFromFile(input);

			// Remove the first text box in the document
			doc.TextBoxes.RemoveAt(0);

			// Clear all the text boxes in the document
			//doc.TextBoxes.Clear();

			// Specify the output file path
			string output = "RemoveTextBox.docx";

			// Save the modified document to the output file with the specified file format (Docx)
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to release resources
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
