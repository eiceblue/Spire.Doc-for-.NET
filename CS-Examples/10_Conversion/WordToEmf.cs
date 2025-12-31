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

namespace WordToEmf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a Word document.
			Document document = new Document();

			//Load the file from disk.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx", FileFormat.Docx);

			//Convert the first page of document to image.
			System.Drawing.Image image = document.SaveToImages(0, Spire.Doc.Documents.ImageType.Metafile);
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

            string result = "Result-WordToEmf.emf";

			//Save the file.
			image.Save(result, ImageFormat.Emf);

			//Dispose the document
			document.Dispose();

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
