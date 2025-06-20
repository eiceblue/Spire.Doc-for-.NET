using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using System.Data;

namespace RepeatRowOnEachPage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
			// Create a new Word document
			Document document = new Document();

			// Add a section to the document
			Section section = document.AddSection();

			// Add a table to the section
			Table table = section.AddTable(true);

			// Set the preferred width of the table to 100%
			PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
			table.PreferredWidth = width;

			// Add a header row to the table
			TableRow row = table.AddRow();
			row.IsHeader = true;
            // Add a cell to the header row
            TableCell cell = row.AddCell();
			cell.SetCellWidth(100, CellWidthType.Percentage);

            for (int i = 0; i < row.Cells.Count; i++)
            {
                row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }

            // Add a paragraph to the cell with text "Row Header 1"
            Paragraph paragraph = cell.AddParagraph();
			paragraph.AppendText("Row Header 1");
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Add another header row to the table
			row = table.AddRow(false, 1);
			row.IsHeader = true;
            for (int i = 0; i < row.Cells.Count; i++)
            {
                row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Ivory;
            }
            row.Height = 30;
			cell = row.Cells[0];
			cell.SetCellWidth(100, CellWidthType.Percentage);
			cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;

			// Add a paragraph to the cell with text "Row Header 2"
			paragraph = cell.AddParagraph();
			paragraph.AppendText("Row Header 2");
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Add rows and cells to the table
			for (int i = 0; i < 70; i++)
			{
				row = table.AddRow(false, 2);
				cell = row.Cells[0];
				cell.SetCellWidth(50, CellWidthType.Percentage);
				cell.AddParagraph().AppendText("Column 1 Text");
				cell = row.Cells[1];
				cell.SetCellWidth(50, CellWidthType.Percentage);
				cell.AddParagraph().AppendText("Column 2 Text");
			}

			// Set background color for alternating rows
			for (int j = 1; j < table.Rows.Count; j++)
			{
				if (j % 2 == 0)
				{
					TableRow row2 = table.Rows[j];
					for (int f = 0; f < row2.Cells.Count; f++)
					{
						row2.Cells[f].CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;         
                    }
				}
			}

			// Save the document to a file
			String result = "RepeatRowOnEachPage_out.docx";
			document.SaveToFile(result, FileFormat.Docx);

			// Dispose of the document object
			document.Dispose();

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
