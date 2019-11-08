using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetTextDirection
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

            //Add the first section
            Section section1 = doc.AddSection();
            //Set text direction for all text in a section
            section1.TextDirection = TextDirection.RightToLeft;

            //Set Font Style and Size
            ParagraphStyle style = new ParagraphStyle(doc);
            style.Name = "FontStyle";
            style.CharacterFormat.FontName = "Arial";
            style.CharacterFormat.FontSize = 15;
            doc.Styles.Add(style);

            //Add two paragraphs and apply the font style
            Paragraph p = section1.AddParagraph();
            p.AppendText("Only Spire.Doc, no Microsoft Office automation");
            p.ApplyStyle(style.Name);
            p = section1.AddParagraph();
            p.AppendText("Convert file documents with high quality");
            p.ApplyStyle(style.Name);

            //Set text direction for a part of text
            //Add the second section
            Section section2 = doc.AddSection();
            //Add a table
            Table table = section2.AddTable();
            table.ResetCells(1, 1);
            TableCell cell = table.Rows[0].Cells[0];
            table.Rows[0].Height = 150;
            table.Rows[0].Cells[0].Width = 10;
            //Set vertical text direction of table
            cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated;
            cell.AddParagraph().AppendText("This is vertical style");
            //Add a paragraph and set horizontal text direction
            p = section2.AddParagraph();
            p.AppendText("This is horizontal style");
            p.ApplyStyle(style.Name);

            //Save and launch document
            string output = "SetTextDirection.docx";
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
