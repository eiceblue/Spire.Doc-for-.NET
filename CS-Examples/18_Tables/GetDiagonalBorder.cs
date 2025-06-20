using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using System.Text;
using System.IO;

namespace GetDiagonalBorder
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\GetDiagonalBorderOfCell.docx");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Get the first table in the section
			Table table = section.Tables[0] as Table;

			// Create a StringBuilder to store the border information
			StringBuilder stringBuilder = new StringBuilder();

			// Get the DiagonalUp border type of cell (0,0) in the table
			Spire.Doc.Documents.BorderStyle bs_UP = table[0, 0].CellFormat.Borders.DiagonalUp.BorderType;
			stringBuilder.AppendLine("DiagonalUp border type of table cell (0,0) is " + bs_UP);

			// Get the DiagonalUp border color of cell (0,0) in the table
			Color color_UP = table[0, 0].CellFormat.Borders.DiagonalUp.Color;
			stringBuilder.AppendLine("DiagonalUp border color of table cell (0,0) is " + color_UP);

			// Get the line width of the DiagonalUp border of cell (0,0) in the table
			float width_UP = table[0, 0].CellFormat.Borders.DiagonalUp.LineWidth;
			stringBuilder.AppendLine("Line width of DiagonalUp border of table cell (0,0) is " + width_UP);

			// Get the DiagonalDown border type of cell (0,0) in the table
			Spire.Doc.Documents.BorderStyle bs_Down = table[0, 0].CellFormat.Borders.DiagonalDown.BorderType;
			stringBuilder.AppendLine("DiagonalDown border type of table cell (0,0) is " + bs_Down);

			// Get the DiagonalDown border color of cell (0,0) in the table
			Color color_Down = table[0, 0].CellFormat.Borders.DiagonalDown.Color;
			stringBuilder.AppendLine("DiagonalDown border color of table cell (0,0) is " + color_Down);

			// Get the line width of the DiagonalDown border of cell (0,0) in the table
			float width_Down = table[0, 0].CellFormat.Borders.DiagonalDown.LineWidth;
			stringBuilder.AppendLine("DiagonalDown border line width of table cell (0,0) is " + width_UP);

			// Specify the output file path
			string output = "GetDiagonalBorder_out.txt";

			// Write the border information to the output file
			File.WriteAllText(output, stringBuilder.ToString());

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launching the Word file.
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
