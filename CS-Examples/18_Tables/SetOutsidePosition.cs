using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetOutsidePosition
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
			Section sec = doc.AddSection();

			// Get the header of the first section in the document
			HeaderFooter header = doc.Sections[0].HeadersFooters.Header;

			// Add a paragraph to the header with left-aligned text
			Paragraph paragraph = header.AddParagraph();
			paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

			// Append an image to the paragraph in the header
			DocPicture headerimage = paragraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));

			// Add a table to the header
			Table table = header.AddTable();
			table.ResetCells(4, 2);

			// Set table properties for text wrapping and positioning
			table.Format.WrapTextAround = true;
			table.Format.Positioning.HorizPositionAbs = HorizontalPosition.Outside;
			table.Format.Positioning.VertRelationTo = VerticalRelation.Margin;
			table.Format.Positioning.VertPosition = 43;

			// Define data for the table cells
			String[][] data = {
				new string[] {"Spire.Doc.left", "Spire XLS.right"},
				new string[] {"Spire.Presentatio.left", "Spire.PDF.right"},
				new string[] {"Spire.DataExport.left", "Spire.PDFViewe.right"},
				new string[] {"Spire.DocViewer.left", "Spire.BarCode.right"}
			};

			// Fill the table with data and set cell widths
			for (int r = 0; r < 4; r++)
			{
				TableRow dataRow = table.Rows[r];
				for (int c = 0; c < 2; c++)
				{
					if (c == 0)
					{
						// Add left-aligned text to the cell
						Paragraph par = dataRow.Cells[c].AddParagraph();
						par.AppendText(data[r][c]);
						par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
						dataRow.Cells[c].SetCellWidth(180, CellWidthType.Point);
					}
					else
					{
						// Add right-aligned text to the cell
						Paragraph par = dataRow.Cells[c].AddParagraph();
						par.AppendText(data[r][c]);
						par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
						dataRow.Cells[c].SetCellWidth(180, CellWidthType.Point);
					}
				}
			}

			// Specify the output file name
			string output = "SetOutsidePosition.docx";

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
