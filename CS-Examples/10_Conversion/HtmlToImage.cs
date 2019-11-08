using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing.Imaging;

namespace HtmlToImage
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_HtmlFile1.html", FileFormat.Html, XHTMLValidationType.None);
            
            String result = "Result-HtmlToImage.png";

            //Save to image. You can convert HTML to BMP, JPEG, PNG, GIF, Tiff£¬etc.
            Image image = document.SaveToImages(0, ImageType.Bitmap);
            image.Save(result, ImageFormat.Png);

            //Launch the image.
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
