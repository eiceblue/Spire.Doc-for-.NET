using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertingTextbox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document and and a section.
            Document document = new Document();
            Section section=document.AddSection();

            InsertTextbox(section);

            //Save docx file.
            document.SaveToFile("Sample.docx",FileFormat.Docx);

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }


        private void InsertTextbox(Section section)
        {
            Paragraph paragraph
                = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();

            //Insert and format the first textbox.
            Spire.Doc.Fields.TextBox textBox1 = paragraph.AppendTextBox(240, 35);
            textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            textBox1.Format.LineColor = System.Drawing.Color.Gray;
            textBox1.Format.LineStyle = TextBoxLineStyle.Simple;
            textBox1.Format.FillColor = System.Drawing.Color.DarkSeaGreen;
            Paragraph para = textBox1.Body.AddParagraph();
            TextRange txtrg = para.AppendText("Textbox 1 in the document");
            txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
            txtrg.CharacterFormat.FontSize = 14;
            txtrg.CharacterFormat.TextColor = System.Drawing.Color.White;
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            //Insert and format the second textbox.
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            Spire.Doc.Fields.TextBox textBox2 = paragraph.AppendTextBox(240, 35);
            textBox2.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            textBox2.Format.LineColor = System.Drawing.Color.Tomato;
            textBox2.Format.LineStyle = TextBoxLineStyle.ThinThick;
            textBox2.Format.FillColor = System.Drawing.Color.Blue;
            textBox2.Format.LineDashing = LineDashing.Dot;
            para = textBox2.Body.AddParagraph();
            txtrg = para.AppendText("Textbox 2 in the document");
            txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
            txtrg.CharacterFormat.FontSize = 14;
            txtrg.CharacterFormat.TextColor = System.Drawing.Color.Pink;
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            //Insert and format the third textbox.
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            Spire.Doc.Fields.TextBox textBox3 = paragraph.AppendTextBox(240, 35);
            textBox3.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            textBox3.Format.LineColor = System.Drawing.Color.Violet;
            textBox3.Format.LineStyle = TextBoxLineStyle.Triple;
            textBox3.Format.FillColor = System.Drawing.Color.Pink;
            textBox3.Format.LineDashing = LineDashing.DashDotDot;
            para = textBox3.Body.AddParagraph();
            txtrg = para.AppendText("Textbox 3 in the document");
            txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
            txtrg.CharacterFormat.FontSize = 14;
            txtrg.CharacterFormat.TextColor = System.Drawing.Color.Tomato;
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
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
