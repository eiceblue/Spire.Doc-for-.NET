using System;
using System.Windows.Forms;
using Spire.Doc;
using System.Text;
using System.IO;

namespace GetRowCellIndex
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Word from disk
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextInTable.docx");

            //Get the first section
            Section section = doc.Sections[0];

            //Get the first table in the section
            Table table = section.Tables[0] as Table;

            StringBuilder content = new StringBuilder();

            //Get table collections
            Spire.Doc.Collections.TableCollection collections = section.Tables;

            //Get the table index
            int tableIndex = collections.IndexOf(table);

            //Get the index of the last table row
            TableRow row = table.LastRow;
            int rowIndex = row.GetRowIndex();

            //Get the index of the last table cell
            TableCell cell = row.LastChild as TableCell;
            int cellIndex = cell.GetCellIndex();

            //Append these information into content
            content.AppendLine("Table index is " + tableIndex.ToString());
            content.AppendLine("Row index is " + rowIndex.ToString());
            content.AppendLine("Cell index is " + cellIndex.ToString());

            //Save to txt file
            string output = "GetRowCellIndex_out.txt";
            File.WriteAllText(output, content.ToString());

            //Launch the file
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
