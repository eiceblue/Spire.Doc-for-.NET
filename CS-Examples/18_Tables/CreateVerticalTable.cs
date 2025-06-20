
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CreateVerticalTable
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

			// Add a new section to the document
			Section section = document.AddSection();

			// Add a table to the section
			Table table = section.AddTable();
			table.ResetCells(1, 1);

			// Get the first cell of the table
			TableCell cell = table.Rows[0].Cells[0];

			// Set the height of the table row
			table.Rows[0].Height = 150;

			// Add a paragraph with text to the cell
			cell.AddParagraph().AppendText("Draft copy in vertical style");

			// Set the text direction of the cell to right-to-left rotated
			cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated;

			// Enable wrap text around the table
			table.Format.WrapTextAround = true;

			// Set the vertical position of the table relative to the page
			table.Format.Positioning.VertRelationTo = VerticalRelation.Page;

			// Set the horizontal position of the table relative to the page
			table.Format.Positioning.HorizRelationTo = HorizontalRelation.Page;

			// Set the horizontal position of the table
			table.Format.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width;

			// Set the vertical position of the table
			table.Format.Positioning.VertPosition = 200;

			// Save the document to a file in Docx2013 format
			String result = "Result-CreateVerticalTable.docx";
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
