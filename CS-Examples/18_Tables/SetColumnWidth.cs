using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace SetColumnWidth
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
            Table table = section.Tables[0] as Table;

            //Traverse the first column
            for (int i = 0; i < table.Rows.Count; i++)
            {
                //Set the cell width type
                table.Rows[i].Cells[0].CellWidthType = CellWidthType.Point;
                //Set the value
                table.Rows[i].Cells[0].Width = 200;
            }

            //Save to file
            document.SaveToFile(@"ColumnWidth.docx", FileFormat.Docx);
            //Launch the document
            FileViewer("ColumnWidth.docx");
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
