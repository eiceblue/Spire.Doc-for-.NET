using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Interface;

namespace DifferentBorders
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\TableSample.docx");

			// Get the first table in the document's first section
			Table table = document.Sections[0].Tables[0] as Table;

			// Set borders for the entire table
			setTableBorders(table);

			// Set borders for a specific cell in the table
			setCellBorders(table.Rows[2].Cells[0]);

			// Save the modified document to a new file
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launch the MS Word file
            WordDocViewer("Sample.docx");
        }

        private void setTableBorders(Table table)
        {
            table.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
            table.Format.Borders.LineWidth = 3.0F;
            table.Format.Borders.Color = Color.Red;
        }

        private void setCellBorders(TableCell tableCell)
        {
            tableCell.CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.DotDash;
            tableCell.CellFormat.Borders.LineWidth = 1.0F;
            tableCell.CellFormat.Borders.Color = Color.Green;
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
