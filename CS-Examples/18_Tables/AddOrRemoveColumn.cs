using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace AddOrRemoveColumn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Load the document from disk.
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_N2.docx");

            //Access the first section
            Section section = doc.Sections[0];

            //Access the first table
            Table table = section.Tables[0] as Table;

            //Add a blank column
            int columnIndex1 = 0;
            AddColumn(table, columnIndex1);

            //Remove a column
            int columnIndex2 = 2;
            RemoveColumn(table, columnIndex2);

            //Save the Word file
            string output = "AddOrRemoveColumn_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

            //Launch the file
            FileViewer(output);
        }
        private void AddColumn(Table table, int columnIndex)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                TableCell addCell = new TableCell(table.Document);
                table.Rows[r].Cells.Insert(columnIndex, addCell);
            }
        }
        private void RemoveColumn(Table table, int columnIndex)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                table.Rows[r].Cells.RemoveAt(columnIndex);
            }
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
