using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetTextDirection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           // Create a new instance of the Document class.
			Document doc = new Document();

			// Add a section to the document.
			Section section1 = doc.AddSection();

			// Set the text direction of section1 to right-to-left.
			section1.TextDirection = TextDirection.RightToLeft;

			// Create a new paragraph style and set its properties.
			ParagraphStyle style = new ParagraphStyle(doc);
			style.Name = "FontStyle";
			style.CharacterFormat.FontName = "Arial";
			style.CharacterFormat.FontSize = 15;

			// Add the style to the document's styles collection.
			doc.Styles.Add(style);

			// Add a paragraph to section1, append text, and apply the created style.
			Paragraph p = section1.AddParagraph();
			p.AppendText("Only Spire.Doc, no Microsoft Office automation");
			p.ApplyStyle(style.Name);

			// Add another paragraph to section1, append text, and apply the created style.
			p = section1.AddParagraph();
			p.AppendText("Convert file documents with high quality");
			p.ApplyStyle(style.Name);


			// Add another section to the document.
			Section section2 = doc.AddSection();

			// Add a table to section2.
			Table table = section2.AddTable();
			table.ResetCells(1, 1);

			// Access the first cell of the table.
			TableCell cell = table.Rows[0].Cells[0];

			// Set the height of the first row of the table to 150 points.
			table.Rows[0].Height = 150;

			// Set the width of the first cell of the table to 10 points.
			table.Rows[0].Cells[0].SetCellWidth(10, CellWidthType.Point);

			// Set the text direction of the cell to right-to-left rotated.
			cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated;

			// Add a paragraph to the cell and append text.
			cell.AddParagraph().AppendText("This is vertical style");

			// Add another paragraph to section2, append text, and apply the created style.
			p = section2.AddParagraph();
			p.AppendText("This is horizontal style");
			p.ApplyStyle(style.Name);

			// Specify the output file name.
			string output = "SetTextDirection.docx";

			// Save the document to a file with the specified output file name and format (Docx).
			doc.SaveToFile(output, FileFormat.Docx);

			// Clean up resources used by the document.
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
