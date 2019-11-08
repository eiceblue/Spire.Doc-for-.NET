using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ImageWaterMark
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
			
			//Insert the imgae watermark.
            InsertImageWatermark(document);
            //Save as docx file.
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }

        private void InsertImageWatermark(Document document)
        {
            PictureWatermark picture = new PictureWatermark();
            picture.Picture = System.Drawing.Image.FromFile(@"..\..\..\..\..\..\Data\ImageWatermark.png");
            picture.Scaling = 250;
            picture.IsWashout = false;
            document.Watermark = picture;
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
