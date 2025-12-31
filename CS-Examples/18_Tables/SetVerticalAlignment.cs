using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetVerticalAlignment
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

			// Add a section to the document
			Section section = doc.AddSection();

			// Add a table to the section with auto-fit behavior
			Table table = section.AddTable(true);

			// Reset the table cells to 3 rows and 3 columns
			table.ResetCells(3, 3);

			// Apply vertical merging to the first column of the table, spanning 3 rows
			table.ApplyVerticalMerge(0, 0, 2);

			// Set the vertical alignment of cells in the table
			table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
			table.Rows[0].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Top;
			table.Rows[0].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Top;
			table.Rows[1].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
			table.Rows[1].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
			table.Rows[2].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;
			table.Rows[2].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;

			// Add a paragraph to the first cell of the first row, and append an image to it
			Paragraph paraPic = table.Rows[0].Cells[0].AddParagraph();
			DocPicture pic = paraPic.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture pic = paraPic.AppendPicture(@"..\..\..\..\..\..\Data\E-iceblue.png");
            */

            // Define data for the table cells
            String[][] data = {
				new string[] {"", "Spire.Office", "Spire.DataExport"},
				new string[] {"", "Spire.Doc", "Spire.DocViewer"},
				new string[] {"", "Spire.XLS", "Spire.PDF"}
			};

			// Fill the table with data and set cell widths
			for (int r = 0; r < 3; r++)
			{
				TableRow dataRow = table.Rows[r];
				dataRow.Height = 50;
				for (int c = 0; c < 3; c++)
				{
					if (c == 1)
					{
						// Add text to the cell and set its width
						Paragraph par = dataRow.Cells[c].AddParagraph();
						par.AppendText(data[r][c]);
						dataRow.Cells[c].SetCellWidth((section.PageSetup.ClientWidth) / 2, CellWidthType.Point);
					}
					if (c == 2)
					{
						// Add text to the cell and set its width
						Paragraph par = dataRow.Cells[c].AddParagraph();
						par.AppendText(data[r][c]);
						dataRow.Cells[c].SetCellWidth((section.PageSetup.ClientWidth) / 2, CellWidthType.Point);
					}
				}
			}

			// Specify the output file name
			string output = "SetVerticalAlignment.docx";

			// Save the document to a file in Docx format
			doc.SaveToFile(output, FileFormat.Docx);

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
