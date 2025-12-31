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

namespace ConvertToImage
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

			//Load the file from disk.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertedTemplate.docx");

			//Save the first page to image
			Image img = document.SaveToImages(0, ImageType.Bitmap);

			//Save the image to file
			img.Save("sample.png", ImageFormat.Png);

            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            SkiaSharp.SKImage images = document.SaveToImages(0, ImageType.Bitmap);
            FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);
            images.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100).SaveTo(fileStream);
            fileStream.Flush();
            */

            //////////////////Use the following code for WPF dlls/////////////////////////
            /*
            BitmapSource[] images = document.SaveToImages(Spire.Doc.Documents.ImageType.Bitmap);
            PngBitmapEncoder pE = new PngBitmapEncoder();
            pE.Frames.Add(BitmapFrame.Create(images[0]));
            string outputfile = String.Format(outputFile, ImageFormat.Png);
            using (Stream stream = File.Create(outputfile))
            {
                pE.Save(stream);
            }
            */


            //Dispose the document
            document.Dispose();

            //Launching the image file.
            WordDocViewer("sample.png");
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
