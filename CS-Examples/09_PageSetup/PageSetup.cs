using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace PageSetup
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
			Document document = new Document();

			// Add a new section to the document.
			Section section = document.AddSection();

			// Set the page size of the section to A4.
			section.PageSetup.PageSize = PageSize.A4;

			// Set the top margin of the section to 72 points (1 inch).
			section.PageSetup.Margins.Top = 72f;

			// Set the bottom margin of the section to 72 points (1 inch).
			section.PageSetup.Margins.Bottom = 72f;

			// Set the left margin of the section to 89.85 points (approximately 1.27 cm).
			section.PageSetup.Margins.Left = 89.85f;

			// Set the right margin of the section to 89.85 points (approximately 1.27 cm).
			section.PageSetup.Margins.Right = 89.85f;

			// Call a method to insert headers and footers in the section.
			InsertHeaderAndFooter(section);

			// Call a method to add a table to the section.
			addTable(section);

			// Save the document to a file in the Doc format (older Word format).
			document.SaveToFile("Sample.doc", FileFormat.Doc);

			// Release the resources associated with the document.
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("Sample.doc");


        }

        private void addTable(Section section)
		{
			// Array for table header.
			String[] header = { "Name", "Capital", "Continent", "Area", "Population" };

			// 2D Array for table data.
			String[][] data =
                {
                    new String[]{"Argentina", "Buenos Aires", "South America", "2777815", "32300003"},
                    new String[]{"Bolivia", "La Paz", "South", "1098575", "7300000"},
                    new String[]{"Brazil", "Brasilia", "South", "8511196", "150400000"},
                    new String[]{"Canada", "Ottawa", "North", "9976147", "26500000"},
                    new String[]{"Chile", "Santiago", "South", "756943", "13200000"},
                    new String[]{"Colombia", "Bagota", "South", "1138907", "33000000"},
                    new String[]{"Cuba", "Havana", "North", "114524", "10600000"},
                    new String[]{"Ecuador", "Quito", "South", "455502", "10600000"},
                    new String[]{"El Salvador", "San Salvador", "North", "20865", "5300000"},
                    new String[]{"Guyana", "Georgetown", "South", "214969", "800000"},
                    new String[]{"Jamaica", "Kingston", "North", "11424", "2500000"},
                    new String[]{"Mexico", "Mexico City", "North", "1967180", "88600000"},
                    new String[]{"Nicaragua", "Managua", "North", "139000", "3900000"},
                    new String[]{"Paraguay", "Asuncion", "South", "406576", "4660000"},
                    new String[]{"Peru", "Lima", "South", "1285215", "21600000"},
                    new String[]{"United States", "Washington", "North", "9363130", "249200000"},
                    new String[]{"Uruguay", "Montevideo", "South", "176140", "3002000"},
                    new String[]{"Venezuela", "Caracas", "South", "912047", "19700000"}
                };
			
			// Add a table to the section and enable autofit.
			Spire.Doc.Table table = section.AddTable(true);
			
			// Set the number of rows and columns in the table.
			table.ResetCells(data.Length + 1, header.Length);

			// First Row (Table Header)
			TableRow row = table.Rows[0];
			row.IsHeader = true;
			row.Height = 20;
			row.HeightType = TableRowHeightType.Exactly;
            for (int i = 0; i < row.Cells.Count; i++)
            {
                row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Gray;
            }

            // Populate the header cells with text and formatting.
            for (int i = 0; i < header.Length; i++)
			{
				row.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
				Paragraph p = row.Cells[i].AddParagraph();
				p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
				TextRange txtRange = p.AppendText(header[i]);
				txtRange.CharacterFormat.Bold = true;
			}

			// Data Rows
			for (int r = 0; r < data.Length; r++)
			{
				TableRow dataRow = table.Rows[r + 1];
				dataRow.Height = 20;
				dataRow.HeightType = TableRowHeightType.Exactly;
                for (int i = 0; i < dataRow.Cells.Count; i++)
                {
                    dataRow.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Empty;
                }

                // Populate the data cells with text.
                for (int c = 0; c < data[r].Length; c++)
				{
					dataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
					dataRow.Cells[c].AddParagraph().AppendText(data[r][c]);
				}
			}
		}

		private void InsertHeaderAndFooter(Section section)
		{
			// Get the header and footer objects from the section.
			HeaderFooter header = section.HeadersFooters.Header;
			HeaderFooter footer = section.HeadersFooters.Footer;

			// Add a paragraph to the header and insert an image and text.
			Paragraph headerParagraph = header.AddParagraph();
			DocPicture headerPicture = headerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Header.png"));
			TextRange text = headerParagraph.AppendText("Demo of Spire.Doc");
			text.CharacterFormat.FontName = "Arial";
			text.CharacterFormat.FontSize = 10;
			text.CharacterFormat.Italic = true;
            headerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
            headerParagraph.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;
			headerParagraph.Format.Borders.Bottom.Space = 0.05F;

			// Set picture properties for the header image.
			headerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
			headerPicture.HorizontalOrigin = HorizontalOrigin.Page;
			headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			headerPicture.VerticalOrigin = VerticalOrigin.Page;
			headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top;

			// Add a paragraph to the footer and insert an image and fields for page numbering.
			Paragraph footerParagraph = footer.AddParagraph();
			DocPicture footerPicture = footerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Footer.png"));
			footerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
			footerPicture.HorizontalOrigin = HorizontalOrigin.Page;
			footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			footerPicture.VerticalOrigin = VerticalOrigin.Page;
			footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom;

			// Insert fields for page numbering.
			footerParagraph.AppendField("page number", FieldType.FieldPage);
			footerParagraph.AppendText(" of ");
			footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
            footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
            footerParagraph.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
			footerParagraph.Format.Borders.Top.Space = 0.05F;
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
