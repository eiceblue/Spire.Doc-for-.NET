using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetVerticalAlignment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new Word document and add a new section
            Document doc = new Document();
            Section section = doc.AddSection();

            //Add a table with 3 columns and 3 rows
            Table table = section.AddTable(true);
            table.ResetCells(3, 3);

            //Merge rows
            table.ApplyVerticalMerge(0, 0, 2);

            //Set the vertical alignment for each cell, default is top
            table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            table.Rows[0].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Top;
            table.Rows[0].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Top;
            table.Rows[1].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            table.Rows[1].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            table.Rows[2].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;
            table.Rows[2].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;

            //Inset a picture into the table cell
            Paragraph paraPic = table.Rows[0].Cells[0].AddParagraph();
            DocPicture pic = paraPic.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));

            //Create data
            String[][] data = {
                       new string[] {"","Spire.Office","Spire.DataExport"},
                       new string[] {"","Spire.Doc","Spire.DocViewer"},
                       new string[] {"","Spire.XLS","Spire.PDF"}
                            };

            //Append data to table
            for (int r = 0; r < 3; r++)
            {
                TableRow dataRow = table.Rows[r];
                dataRow.Height = 50;
                for (int c = 0; c < 3; c++)
                {
                    if (c == 1)
                    {
                        Paragraph par = dataRow.Cells[c].AddParagraph();
                        par.AppendText(data[r][c]);
                        dataRow.Cells[c].Width = (section.PageSetup.ClientWidth) / 2;
                    }
                    if (c == 2)
                    {
                        Paragraph par = dataRow.Cells[c].AddParagraph();
                        par.AppendText(data[r][c]);
                        dataRow.Cells[c].Width = (section.PageSetup.ClientWidth) / 2;
                    }
                }
            }

            //Save and launch document
            string output = "SetVerticalAlignment.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
