using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace HelloWorld
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
            Document document = new Document();

            //Create a new section
            Section section = document.AddSection();

            //Create a new paragraph
            Paragraph paragraph = section.AddParagraph();

            //Append Text
            paragraph.AppendText("Hello World!");

            //Save doc file.
            document.SaveToFile("Sample.docx",FileFormat.Docx);

            //Launching the Word file.
            WordDocViewer("Sample.docx");


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
