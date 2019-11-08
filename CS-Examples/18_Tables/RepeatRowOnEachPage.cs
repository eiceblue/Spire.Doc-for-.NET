using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

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
            //Create word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();

            //Create a table width default borders
            Table table = section.AddTable(true);
            //Set table with to 100%
            PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
            table.PreferredWidth = width;

            //Add a new row 
            TableRow row = table.AddRow();
            //Set the row as a table header 
            row.IsHeader = true;
            //Set the backcolor of row
            row.RowFormat.BackColor = Color.LightGray;
            //Add a new cell for row
            TableCell cell = row.AddCell();
            cell.SetCellWidth(100, CellWidthType.Percentage);
            //Add a paragraph for cell to put some data
            Paragraph parapraph = cell.AddParagraph();
            //Add text 
            parapraph.AppendText("Row Header 1");
            //Set paragraph horizontal center alignment
            parapraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            row = table.AddRow(false, 1);
            row.IsHeader = true;
            row.RowFormat.BackColor = Color.Ivory;
            //Set row height
            row.Height = 30;
            cell = row.Cells[0];
            cell.SetCellWidth(100, CellWidthType.Percentage);
            //Set cell vertical middle alignment
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            //Add a paragraph for cell to put some data
            parapraph = cell.AddParagraph();
            //Add text 
            parapraph.AppendText("Row Header 2");
            parapraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            //Add many common rows 
            for (int i = 0; i < 70; i++)
            {
                row = table.AddRow(false,2);
                cell = row.Cells[0];
                //Set cell width to 50% of table width
                cell.SetCellWidth(50, CellWidthType.Percentage);
                cell.AddParagraph().AppendText("Column 1 Text");
                cell = row.Cells[1];
                cell.SetCellWidth(50, CellWidthType.Percentage);
                cell.AddParagraph().AppendText("Column 2 Text");
            }
            //Set cell backcolor
            for (int j = 1; j < table.Rows.Count; j++)
            {
                if (j % 2 == 0)
                {
                    TableRow row2 = table.Rows[j];
                    for (int f = 0; f < row2.Cells.Count; f++)
                    {
                        row2.Cells[f].CellFormat.BackColor = Color.LightBlue;
                    }
                }
               
            }

            String result = "RepeatRowOnEachPage_out.docx";

            //Save file.
            document.SaveToFile(result,FileFormat.Docx);

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
