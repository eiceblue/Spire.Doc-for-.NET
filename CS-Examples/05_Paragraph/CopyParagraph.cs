using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CopyParagraph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document1.
            Document document1 = new Document();

            //Load the file from disk.
            document1.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_5.docx");

            //Create a new document.
            Document document2 = new Document();

            //Get paragraph 1 and paragraph 2 in document1.
            Section s = document1.Sections[0];
            Paragraph p1 = s.Paragraphs[0];
            Paragraph p2 = s.Paragraphs[1];

            //Copy p1 and p2 to document2.
            Section s2 = document2.AddSection();
            Paragraph NewPara1 = (Paragraph)p1.Clone();
            s2.Paragraphs.Add(NewPara1);
            Paragraph NewPara2 = (Paragraph)p2.Clone();
            s2.Paragraphs.Add(NewPara2);

            //Add watermark.
            PictureWatermark WM = new PictureWatermark();
            WM.Picture = Image.FromFile(@"..\..\..\..\..\..\Data\Logo.jpg");
            document2.Watermark = WM;

            String result = "Result-CopyWordParagraph.docx";

            //Save the file.
            document2.SaveToFile(result, FileFormat.Docx2013);

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
