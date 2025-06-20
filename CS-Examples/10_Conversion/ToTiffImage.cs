using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ToTiffImage
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertedTemplate.docx");
			
			//Save the document to a tiff image.
			JoinTiffImages(SaveAsImage(document),"Sample.tif",EncoderValue.CompressionLZW);

			//Dispose the document
			document.Dispose();
			
            //Launching the tiff file.
            FileViewer("Sample.tif");
        }

		private static Image[] SaveAsImage(Document document)
        {    
			//Save all the pages in the document to images.
            Image[] images = document.SaveToImages(ImageType.Bitmap);    
            return images;
        }

        private static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
			//Set the code information for the images.
            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
            for (int j = 0; j < encoders.Length; j++)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            throw new Exception(mimeType + " mime type not found in ImageCodecInfo");
        }

        public static void JoinTiffImages(Image[] images, string outFile, EncoderValue compressEncoder)
        {
			//Set the encoder parameters.
            System.Drawing.Imaging.Encoder enc = System.Drawing.Imaging.Encoder.SaveFlag;
            EncoderParameters ep = new EncoderParameters(2);
            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
            ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)compressEncoder);
            Image pages = images[0];
            int frame = 0;
            ImageCodecInfo info = GetEncoderInfo("image/tiff");
            foreach (Image img in images)
            {
                if (frame == 0)
                {
                    pages = img;
                    //Save the first frame.
                    pages.Save(outFile, info, ep);
                }

                else
                {
                    //Save the intermediate frames.
                    ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);

                    pages.SaveAdd(img, ep);
                }
                if (frame == images.Length - 1)
                {
                    //Flush and close.
                    ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.Flush);
                    pages.SaveAdd(ep);
                }
                frame++;
            }
        }
		
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
