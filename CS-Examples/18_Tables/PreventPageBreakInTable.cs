using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace PreventPageBreakInTable
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
			Document document = new Document();

			// Load an existing Word document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_5.docx");

			// Get the first table in the first section of the document
			Table table = document.Sections[0].Tables[0] as Table;

			// Iterate through each row in the table
			foreach (TableRow row in table.Rows)
			{
				// Iterate through each cell in the row
				foreach (TableCell cell in row.Cells)
				{
					// Iterate through each paragraph in the cell
					foreach (Paragraph p in cell.Paragraphs)
					{
						// Set "Keep with next" property to true to prevent page breaks within paragraphs
						p.Format.KeepFollow = true;
					}
				}
			}

			// Specify the output file path
			String result = "Result-PreventPageBreaksInWordTable.docx";

			// Save the modified document to a file
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launch the MS Word file.
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
