using System;
using System.Data;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ReplaceWithTable
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
            string input = @"..\..\..\..\..\..\Data\Bookmark.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Create a table
            Table table = new Table(doc, true);

            //Create datatable
            DataTable dt = new DataTable();
            dt.Columns.Add("id", typeof(string));
            dt.Columns.Add("name", typeof(string));
            dt.Columns.Add("job", typeof(string));
            dt.Columns.Add("email", typeof(string));
            dt.Columns.Add("salary", typeof(string));
            dt.Rows.Add(new string[] { "Name", "Capital", "Continent", "Area", "Population" });
            dt.Rows.Add(new string[] { "Argentina", "Buenos Aires", "South America", "2777815", "32300003" });
            dt.Rows.Add(new string[] { "Bolivia", "La Paz", "South America", "1098575", "7300000" });
            dt.Rows.Add(new string[] { "Brazil", "Brasilia", "South America", "8511196", "150400000" });
            table.ResetCells(dt.Rows.Count, dt.Columns.Count);

            //Fill the table with the data of datatable
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    table.Rows[i].Cells[j].AddParagraph().AppendText(dt.Rows[i][j].ToString());
                }
            }

            //Get the specific bookmark by its name
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark("Test");

            //Create a TextBodyPart instance and add the table to it
            TextBodyPart part = new TextBodyPart(doc);
            part.BodyItems.Add(table);

            //Replace the current bookmark content with the TextBodyPart object
            navigator.ReplaceBookmarkContent(part);

            //Save and launch document
            string output = "ReplaceWithTable.docx";
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
