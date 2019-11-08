using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace ReplaceTextInTable
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

            //Define a regular expression to match the {} with its content
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"{[^\}]+\}");

            //Replace the text of table with regex
            table.Replace(regex, "E-iceblue");

            //Replace old text with new text in table
            table.Replace("Beijing", "Component", false, true);

            //Save the Word document
            string output="ReplaceTextInTable_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

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
