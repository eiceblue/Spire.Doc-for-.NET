using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ReadTableFromTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        
			// Load the Word document from the specified file path
			string input = @"..\..\..\..\..\..\Data\TextBoxTable.docx";
			Document doc = new Document();
			doc.LoadFromFile(input);

			// Get the first textbox in the document
			Spire.Doc.Fields.TextBox textbox = doc.TextBoxes[0];

			// Get the first table from the textbox
			Table table = textbox.Body.Tables[0] as Table;

			// Initialize an empty string to store the table data
			string str = null;

			// Iterate through each row in the table
			foreach (TableRow row in table.Rows)
			{
				// Iterate through each cell in the row
				foreach (TableCell cell in row.Cells)
				{
					// Iterate through each paragraph in the cell
					foreach (Paragraph paragraph in cell.Paragraphs)
					{
						// Append the text of each paragraph to the string, separated by a tab
						str += paragraph.Text + "\t";
					}
				}
				// Add a new line after processing each row
				str += "\r\n";
			}

			// Specify the output file path
			string output = "ReadTableFromTextBox.txt";

			// Write the table data to the output file
			File.WriteAllText(output, str);

			// Dispose of the document object
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
