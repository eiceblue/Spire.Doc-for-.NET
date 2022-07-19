using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Fields.Shape;

namespace SetSpaceBetweenAsianAndLatinText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {//Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\SetSpaceBetweenAsianAndLatinText.docx");

            Paragraph para = document.Sections[0].Paragraphs[0];

            //Set whether to automatically adjust space between Asian text and Latin text
            para.Format.AutoSpaceDE = false;
            //Set whether to automatically adjust space between Asian text and numbers
            para.Format.AutoSpaceDN = true;

            String result = "Result.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launching the MS Word file.
            WordDocViewer(result);
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
