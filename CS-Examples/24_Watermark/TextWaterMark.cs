using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace TextWaterMark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Open a Word document as template.
            Document document = new Document(@"..\..\..\..\..\..\Data\Template.docx");
			
			//Insert text watermark.
            InsertTextWatermark(document.Sections[0]);
            //Save as docx file.
            document.SaveToFile("Sample.docx",FileFormat.Docx);

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }
        private void InsertTextWatermark(Section section)
        {
            TextWatermark txtWatermark = new TextWatermark();
            txtWatermark.Text = "E-iceblue";
            txtWatermark.FontSize = 95;
            txtWatermark.Color = Color.Blue;
            txtWatermark.Layout = WatermarkLayout.Diagonal;
            section.Document.Watermark = txtWatermark;
         
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
