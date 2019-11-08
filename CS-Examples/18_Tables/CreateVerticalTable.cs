
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CreateVerticalTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Add a new section.
            Section section = document.AddSection();

            //Add a table with rows and columns and set the text for the table.
            Table table = section.AddTable();
            table.ResetCells(1, 1);
            TableCell cell = table.Rows[0].Cells[0];
            table.Rows[0].Height = 150;
            cell.AddParagraph().AppendText("Draft copy in vertical style");

            //Set the TextDirection for the table to RightToLeftRotated.
            cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated;

            //Set the table format.
            table.TableFormat.WrapTextAround = true;
            table.TableFormat.Positioning.VertRelationTo = VerticalRelation.Page;
            table.TableFormat.Positioning.HorizRelationTo = HorizontalRelation.Page;
            table.TableFormat.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width;
            table.TableFormat.Positioning.VertPosition = 200;

            String result = "Result-CreateVerticalTable.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
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
