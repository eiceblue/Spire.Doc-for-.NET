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

namespace SetSpacing
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            //Add the text strings to the paragraph and set the style.
            Paragraph para = new Paragraph(document);
            TextRange textRange1 = para.AppendText("This is an inserted paragraph.");
            textRange1.CharacterFormat.TextColor = Color.Blue;
            textRange1.CharacterFormat.FontSize = 15;

            //set the spacing before and after.
            para.Format.BeforeAutoSpacing = false;
            para.Format.BeforeSpacing = 10;
            para.Format.AfterAutoSpacing = false;
            para.Format.AfterSpacing = 10;

            //insert the added paragraph to the word document.
            document.Sections[0].Paragraphs.Insert(1, para);

            String result = "Result-SetTheSpacing.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the file.
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
