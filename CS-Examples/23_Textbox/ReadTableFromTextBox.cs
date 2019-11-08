using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ReadTableFromTextBox
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
            string input = @"..\..\..\..\..\..\Data\TextBoxTable.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first textbox
            Spire.Doc.Fields.TextBox textbox = doc.TextBoxes[0];

            //Get the first table in the textbox
            Table table = textbox.Body.Tables[0] as Table;

            string str = null;

            //Loop through the paragraphs of the table cells and extract them to a .txt file
            foreach (TableRow row in table.Rows)
            {
                foreach (TableCell cell in row.Cells)
                {
                    foreach (Paragraph paragraph in cell.Paragraphs)
                    {
                        str += paragraph.Text + "\t";
                    }
                }
                str += "\r\n";
            }

            //Save to TXT file and launch it
            string output = "ReadTableFromTextBox.txt";
            File.WriteAllText(output, str);
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
