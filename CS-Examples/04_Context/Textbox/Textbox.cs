using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

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
            //Open a blank word document as template
            Document document = new Document(@"..\..\..\..\..\..\Data\Blank.doc");

            InsertTextbox(document.Sections[0]);

            //Save doc file.
            document.SaveToFile("Sample.doc",FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");


        }

        private void InsertTextbox(Section section)
        {
            Paragraph paragraph
                = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
            paragraph.AppendText("The sample demonstrates how to insert a textbox into a document.");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);

            paragraph = section.AddParagraph();
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
            Spire.Doc.Fields.TextBox textBox = paragraph.AppendTextBox(50,20);

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
