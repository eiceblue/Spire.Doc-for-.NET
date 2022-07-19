using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AllowLatinTextWrapInMiddleOfAWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\AllowLatinTextWrapInMiddleOfAWord.docx");

            Paragraph para = document.Sections[0].Paragraphs[0];

            //Allow Latin text to wrap in the middle of a word
            para.Format.WordWrap = false;

            String result = "AllowLatinTextWrapInMiddleOfAWord-Result.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);
          
            //Launching the Word file.
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
