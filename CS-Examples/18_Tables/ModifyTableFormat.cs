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
            //Load Word document from disk
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ModifyTableFormat.docx");

            //Get the first section
            Section section = document.Sections[0];

            //Get tables 
            Table tb1 = section.Tables[0] as Table;
            Table tb2 = section.Tables[1] as Table;
            Table tb3 = section.Tables[2] as Table;

            MoidyTableFormat(tb1);
            ModifyRowFormat(tb2);
            ModifyCellFormat(tb3);

            string output = "ModifyTableFormat_out.docx";
            document.SaveToFile(output, FileFormat.Docx2013);

            //Launch Word file.
            WordDocViewer(output);
        }
        private static void MoidyTableFormat(Table table)
        {
            //Set table width
            table.PreferredWidth=new PreferredWidth(WidthType.Twip,(short)6000);

            //Apply style for table
            table.ApplyStyle(DefaultTableStyle.ColorfulGridAccent3);

            //Set table padding
            table.TableFormat.Paddings.All =5;

            //Set table title and description
            table.Title = "Spire.Doc for .NET";
            table.TableDescription = "Spire.Doc for .NET is a professional Word .NET library";
        }
        private static void ModifyRowFormat(Table table)
        {
            //Set cell spacing
            table.Rows[0].RowFormat.CellSpacing = 2;

            //Set row height
            table.Rows[1].HeightType = TableRowHeightType.Exactly;
            table.Rows[1].Height = 20f;
            
            //Set background color
            table.Rows[2].RowFormat.BackColor = Color.DarkSeaGreen;
         }
        private static void ModifyCellFormat(Table table)
        {
            //Set alignment
            table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            //Set background color
            table.Rows[1].Cells[0].CellFormat.BackColor = Color.DarkSeaGreen;

            //Set cell border
            table.Rows[2].Cells[0].CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
            table.Rows[2].Cells[0].CellFormat.Borders.LineWidth = 1f;
            table.Rows[2].Cells[0].CellFormat.Borders.Left.Color = Color.Red;
            table.Rows[2].Cells[0].CellFormat.Borders.Right.Color = Color.Red;
            table.Rows[2].Cells[0].CellFormat.Borders.Top.Color = Color.Red;
            table.Rows[2].Cells[0].CellFormat.Borders.Bottom.Color = Color.Red;

            //Set text direction
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