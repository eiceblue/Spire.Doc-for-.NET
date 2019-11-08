using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CloneTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the document
            string input = @"..\..\..\..\..\..\Data\TableTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first section
            Section se = doc.Sections[0];

            //Get the first table
            Table original_Table = (Table)se.Tables[0];

            //Copy the existing table to copied_Table via Table.clone()
            Table copied_Table = original_Table.Clone();
            string[] st = new string[] { "Spire.Presentation for .Net", "A professional " +
                "PowerPoint® compatible library that enables developers to create, read, " +
                "write, modify, convert and Print PowerPoint documents on any .NET framework, " +
                ".NET Core platform." };
            //Get the last row of table
            TableRow lastRow = copied_Table.Rows[copied_Table.Rows.Count - 1];
            //Change last row data
            for (int i = 0; i < lastRow.Cells.Count - 1; i++)
            {
                lastRow.Cells[i].Paragraphs[0].Text = st[i];
            }
            //Add copied_Table in section
            se.Tables.Add(copied_Table);

            //Save and launch document
            string output = "CloneTable.docx";
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
