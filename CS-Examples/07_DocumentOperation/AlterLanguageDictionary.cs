using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AlterLanguageDictionary
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

            //Add new section and paragraph to the document.
            Section sec = document.AddSection();
            Paragraph para = sec.AddParagraph();

            //Add a textRange for the paragraph and append some Peru Spanish words.
            TextRange txtRange = para.AppendText("corrige seg¨²n diccionario en ingl¨¦s");
            txtRange.CharacterFormat.LocaleIdASCII = 10250;

            String result = "Result-AlterLanguageDictionary.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
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
