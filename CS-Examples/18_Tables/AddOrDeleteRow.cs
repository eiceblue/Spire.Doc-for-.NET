using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddOrDeleteRow
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a document
            Document document = new Document();
            //Load file
            document.LoadFromFile(@"..\..\..\..\..\..\Data\TableSample.docx");
            Section section = document.Sections[0];
            Table table = section.Tables[0] as Table;

            //Delete the seventh row
            table.Rows.RemoveAt(7);

            //Add a row and insert it into specific position
            TableRow row = new TableRow(document);
            for (int i = 0; i < table.Rows[0].Cells.Count;i++ )
            {
                TableCell tc = row.AddCell();
                Paragraph paragraph = tc.AddParagraph();
                paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                paragraph.AppendText("Added");
            }
            table.Rows.Insert(2,row);
            //Add a row at the end of table
            table.AddRow();

            //Save to file and launch it
            document.SaveToFile("AddDeleteRow.docx", FileFormat.Docx);
            FileViewer("AddDeleteRow.docx");
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
