using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateTableDirectly
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

			// Add a new section to the document
			Section section = doc.AddSection();

			// Create a new table with the document as its parent
			Table table = new Table(doc);
			table.ResetCells(1, 2);

            // Set the preferred width of the table to 100% of the page width
            table.PreferredWidth = new PreferredWidth(WidthType.Percentage, (short)100);

			// Set the border type of the table to single line
            table.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;

			// Create a new row for the table
			TableRow row = table.Rows[0];
  
			// Set the height of the row to 50.0f
			row.Height = 50.0f; 

			// Create the first cell of the row
			TableCell cell1 = table.Rows[0].Cells[0];
			Paragraph para1 = cell1.AddParagraph();
			// Add text to the cell
			para1.AppendText("Row 1, Cell 1"); 
			// Set the horizontal alignment of the text
			para1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center; 
			// Set the background color of the cell
            cell1.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue;
            // Set the vertical alignment of the content in the cell
            cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle; 

			// Create the second cell of the row
			TableCell cell2 = table.Rows[0].Cells[1];
            Paragraph para2 = cell2.AddParagraph();
			// Add text to the cell
			para2.AppendText("Row 1, Cell 2"); 
			// Set the horizontal alignment of the text
			para2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            // Set the background color of the cell
            cell2.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue; 
			// Set the vertical alignment of the content in the cell
			cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle; 

			// Add the table to the section
			section.Tables.Add(table);

			// Save the document to a file in Docx2013 format
			string output = "CreateTableDirectly_out.docx";
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose of the document object to free up resources
			doc.Dispose();

            //Launch the document
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
