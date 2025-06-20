using System;
using System.Windows.Forms;
using Spire.Doc;
using System.Text;
using System.IO;

namespace GetRowCellIndex
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Create a new Document object
			Document doc = new Document();

			// Load an existing Word document from a file
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextInTable.docx");

			// Get the first section of the document
			Section section = doc.Sections[0];

			// Get the first table in the section
			Table table = section.Tables[0] as Table;

			// Create a StringBuilder to store the output content
			StringBuilder content = new StringBuilder();

			// Get the collection of tables in the section
			Spire.Doc.Collections.TableCollection collections = section.Tables;

			// Get the index of the table in the collection
			int tableIndex = collections.IndexOf(table);

			// Get the last row in the table and its index
			TableRow row = table.LastRow;
			int rowIndex = row.GetRowIndex();

			// Get the last cell in the row and its index
			TableCell cell = row.LastChild as TableCell;
			int cellIndex = cell.GetCellIndex();

			// Append the table, row, and cell indices to the output content
			content.AppendLine("Table index is " + tableIndex.ToString());
			content.AppendLine("Row index is " + rowIndex.ToString());
			content.AppendLine("Cell index is " + cellIndex.ToString());

			// Specify the output file path
			string output = "GetRowCellIndex_out.txt";

			// Write the output content to the output file
			File.WriteAllText(output, content.ToString());

			// Dispose of the document object to free up resources
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
