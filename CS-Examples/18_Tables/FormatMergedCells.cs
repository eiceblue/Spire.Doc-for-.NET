using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace FormatMergedCells
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

			// Add a table to the section using the AddTable method
			Table table = AddTable(section);

			// Create a new ParagraphStyle and customize its formatting properties
			ParagraphStyle style = new ParagraphStyle(document);
			style.Name = "Style";
			style.CharacterFormat.TextColor = Color.DeepSkyBlue;
			style.CharacterFormat.Italic = true;
			style.CharacterFormat.Bold = true;
			style.CharacterFormat.FontSize = 13;
			document.Styles.Add(style);

			// Apply horizontal merge for the cells in the first row from column index 0 to 1
			table.ApplyHorizontalMerge(0, 0, 1);

			// Apply the custom style to the paragraph in the first cell of the first row
			table.Rows[0].Cells[0].Paragraphs[0].ApplyStyle(style.Name);

			// Set the vertical alignment and horizontal alignment of the first cell in the first row
			table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
			table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Apply vertical merge for the cells in the second row from row index 1 to 3
			table.ApplyVerticalMerge(0, 1, 3);

			// Apply the custom style to the paragraph in the first cell of the second row
			table.Rows[1].Cells[0].Paragraphs[0].ApplyStyle(style.Name);

			// Set the vertical alignment and horizontal alignment of the first cell in the second row
			table.Rows[1].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
			table.Rows[1].Cells[0].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

			// Set the width of the first cell in the second row as a percentage of the table width
			table.Rows[1].Cells[0].SetCellWidth(20, CellWidthType.Percentage);

			// Save the document to a file in Docx format
			string output = "FormatMergedCells.docx";
			document.SaveToFile(output, FileFormat.Docx);

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launching the file
            WordDocViewer(output);

        }

		 private Table AddTable(Section section)
		{
			// Create a new table with 4 rows and 3 columns
			Table table = section.AddTable(true);
			table.ResetCells(4, 3);

			// Create a DataTable with column headers and data
			DataTable dt = new DataTable();
			dt.Columns.Add();
			dt.Columns.Add();
			dt.Columns.Add();
			dt.Rows.Add("Product", "", "Price");
			dt.Rows.Add("Spire.Doc", "Pro Edition", "$799");
			dt.Rows.Add("", "Standard Edition", "$599");
			dt.Rows.Add("", "Free Edition", "$0");

			// Populate the table cells with data from the DataTable
			for (int r = 0; r < dt.Rows.Count; r++)
			{
				TableRow dataRow = table.Rows[r];
				dataRow.Height = 20;
				dataRow.HeightType = TableRowHeightType.Exactly;

            			for (int i = 0; i < dataRow.Cells.Count; i++)
            			{
                			dataRow.Cells[i].CellFormat.Shading.BackgroundPatternColor =Color.Empty;
            			}
				for (int c = 0; c < dataRow.Cells.Count; c++)
				{
					if (!string.IsNullOrEmpty(dt.Rows[r][c].ToString()))
					{
						TextRange range = dataRow.Cells[c].AddParagraph().AppendText(dt.Rows[r][c].ToString());
						range.CharacterFormat.FontName = "Arial";
					}
				}
			}

			return table;
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
