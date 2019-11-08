using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetTableStyleAndBorder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document and load file
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\TableSample.docx");

            Section section = document.Sections[0];

            //Get the first table
            Table table = section.Tables[0] as Table;

            //Apply the table style
            table.ApplyStyle(DefaultTableStyle.ColorfulList);

            //Set right border of table
            table.TableFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
            table.TableFormat.Borders.Right.LineWidth = 1.0F;
            table.TableFormat.Borders.Right.Color = Color.Red;

            //Set top border of table
            table.TableFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
            table.TableFormat.Borders.Top.LineWidth = 1.0F;
            table.TableFormat.Borders.Top.Color = Color.Green;

            //Set left border of table
            table.TableFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
            table.TableFormat.Borders.Left.LineWidth = 1.0F;
            table.TableFormat.Borders.Left.Color = Color.Yellow;

            //Set bottom border is none
            table.TableFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.DotDash;

            //Set vertical and horizontal border 
            table.TableFormat.Borders.Vertical.BorderType = Spire.Doc.Documents.BorderStyle.Dot;
            table.TableFormat.Borders.Horizontal.BorderType = Spire.Doc.Documents.BorderStyle.None;
            table.TableFormat.Borders.Vertical.Color = Color.Orange;

            //Save the file and launch it
            document.SaveToFile("TableStyleAndBorder.docx", FileFormat.Docx);
            FileViewer("TableStyleAndBorder.docx");
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }       
    }
}
