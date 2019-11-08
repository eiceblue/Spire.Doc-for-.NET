using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertTableIntoTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document
            Document doc = new Document();

            //Add a section
            Section section = doc.AddSection();

            //Add a paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            //Add a textbox to the paragraph
            Spire.Doc.Fields.TextBox textbox = paragraph.AppendTextBox(300, 100);

            //Set the position of the textbox
            textbox.Format.HorizontalOrigin = HorizontalOrigin.Page;
            textbox.Format.HorizontalPosition = 140;
            textbox.Format.VerticalOrigin = VerticalOrigin.Page;
            textbox.Format.VerticalPosition = 50;

            //Add text to the textbox
            Paragraph textboxParagraph = textbox.Body.AddParagraph();
            TextRange textboxRange = textboxParagraph.AppendText("Table 1");
            textboxRange.CharacterFormat.FontName = "Arial";

            //Insert table to the textbox
            Table table = textbox.Body.AddTable(true);

            //Specify the number of rows and columns of the table
            table.ResetCells(4, 4);

            string[,] data = new string[,]
            {
                {"Name","Age","Gender","ID" },
                {"John","28","Male","0023" },
                {"Steve","30","Male","0024" },
                {"Lucy","26","female","0025" }
            };

            //Add data to the table 
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    TextRange tableRange = table[i, j].AddParagraph().AppendText(data[i, j]);
                    tableRange.CharacterFormat.FontName = "Arial";
                }
            }

            //Apply style to the table
            table.ApplyStyle(DefaultTableStyle.TableColorful2);

            //Save and launch document
            string output = "InsertTableIntoTextBox.docx";
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
