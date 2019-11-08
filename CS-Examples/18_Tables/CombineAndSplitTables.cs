using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace CombineAndSplitTables
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Combine tables
            CombineTables();

            //Split a table
            SplitTable();
        }
        private void CombineTables()
        {
            //Load document from disk
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\CombineAndSplitTables.docx");

            //Get the first section
            Section section = doc.Sections[0];

            //Get the first and second table
            Table table1 = section.Tables[0] as Table;
            Table table2 = section.Tables[1] as Table;

            //Add the rows of table2 to table1
            for (int i = 0; i < table2.Rows.Count; i++)
            {
                table1.Rows.Add(table2.Rows[i].Clone());
            }

            //Remove the table2
            section.Tables.Remove(table2);

            //Save the Word file
            string output = "CombineTables_out.docx";
            section.Document.SaveToFile(output, FileFormat.Docx2013);

            //Launch the file
            WordDocViewer(output);

        }
        private void SplitTable()
        {
            //Load document from disk
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\CombineAndSplitTables.docx");

            //Get the first section
            Section section = doc.Sections[0];

            //Get the first table
            Table table = section.Tables[0] as Table;

            //We will split the table at the third row;
            int splitIndex = 2;

            //Create a new table for the split table
            Table newTable = new Table(section.Document);

            //Add rows to the new table
            for (int i = splitIndex; i < table.Rows.Count; i++)
            {
                newTable.Rows.Add(table.Rows[i].Clone());
            }

            //Remove rows from original table
            for (int i = table.Rows.Count - 1; i >= splitIndex; i--)
            {
                table.Rows.RemoveAt(i);
            }

            //Add the new table in section
            section.Tables.Add(newTable);

            //Save the Word file
            string output = "SplitTable_out.docx";
            section.Document.SaveToFile(output, FileFormat.Docx2013);

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
