using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace ModifyTableFormat
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\ModifyTableFormat.docx");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Get the tables in the section
			Table tb1 = section.Tables[0] as Table;
			Table tb2 = section.Tables[1] as Table;
			Table tb3 = section.Tables[2] as Table;

			// Modify the format of tb1
			MoidfyTableFormat(tb1);

			// Modify the row format of tb2
			ModifyRowFormat(tb2);

			// Modify the cell format of tb3
			ModifyCellFormat(tb3);

			// Specify the output file path
			string output = "ModifyTableFormat_out.docx";

			// Save the modified document to a file
			document.SaveToFile(output, FileFormat.Docx2013);

			// Dispose of the document object to free up resources
			document.Dispose();

            //Launch Word file.
            WordDocViewer(output);
        }
		// Modify the table format
		private static void MoidfyTableFormat(Table table)
		{
			// Set the preferred width of the table
			table.PreferredWidth = new PreferredWidth(WidthType.Twip, (short)6000);

			// Apply a specific table style to the table
			table.ApplyStyle(DefaultTableStyle.ColorfulGridAccent3);

			// Set padding for all cells in the table
			table.Format.Paddings.All = 5;

			// Set the title and description of the table
			table.Title = "Spire.Doc for .NET";
			table.TableDescription = "Spire.Doc for .NET is a professional Word .NET library";
		}

		// Modify the row format
		private static void ModifyRowFormat(Table table)
		{
			// Set the cell spacing of the first row
			table.Format.CellSpacing = 2;

            // Set the height of the second row
            table.Rows[1].HeightType = TableRowHeightType.Exactly;
			table.Rows[1].Height = 20f;

			// Set the background color of the third row
            for (int i = 0; i < table.Rows[2].Cells.Count; i++)
            {
                table.Rows[2].Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.DarkSeaGreen;
            }
        }

		// Modify the cell format
		private static void ModifyCellFormat(Table table)
		{
			// Set the vertical alignment and horizontal alignment of the first cell in the first row
			table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
			table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			// Set the background color of the first cell in the second row
            table.Rows[1].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkSeaGreen;

            // Set borders for the first cell in the third row
            table.Rows[2].Cells[0].CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
			table.Rows[2].Cells[0].CellFormat.Borders.LineWidth = 1f;
			table.Rows[2].Cells[0].CellFormat.Borders.Left.Color = Color.Red;
			table.Rows[2].Cells[0].CellFormat.Borders.Right.Color = Color.Red;
			table.Rows[2].Cells[0].CellFormat.Borders.Top.Color = Color.Red;
			table.Rows[2].Cells[0].CellFormat.Borders.Bottom.Color = Color.Red;

			// Set the text direction of the first cell in the fourth row
			table.Rows[3].Cells[0].CellFormat.TextDirection = TextDirection.RightToLeft;
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