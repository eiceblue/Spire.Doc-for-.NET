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
            //Create a Word document
            Document doc = new Document();

            //Add a section
            Section section = doc.AddSection();

            //Create a table 
            Table table = new Table(doc);
            //Set the width of table
            table.PreferredWidth = new PreferredWidth(WidthType.Percentage, (short)100);
            //Set the border of table
            table.TableFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;

            //Create a table row
            TableRow row = new TableRow(doc);
            row.Height = 50.0f;
            table.Rows.Add(row);

            //Create a table cell
            TableCell cell1 = new TableCell(doc);
            //Add a paragraph
            Paragraph para1 = cell1.AddParagraph();
            //Append text in the paragraph
            para1.AppendText("Row 1, Cell 1");
            //Set the horizontal alignment of paragrah
            para1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            //Set the background color of cell
            cell1.CellFormat.BackColor = Color.CadetBlue;
            //Set the vertical alignment of paragraph
            cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            row.Cells.Add(cell1);

            //Create a table cell
            TableCell cell2 = new TableCell(doc);
            Paragraph para2 = cell2.AddParagraph();
            para2.AppendText("Row 1, Cell 2");
            para2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            cell2.CellFormat.BackColor = Color.CadetBlue;
            cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            row.Cells.Add(cell2);

            //Add the table in the section
            section.Tables.Add(table);

            //Save the document
            string output = "CreateTableDirectly_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

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
