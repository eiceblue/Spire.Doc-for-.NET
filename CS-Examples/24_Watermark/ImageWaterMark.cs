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
           
			// Load the document from a template file
			Document document = new Document(@"..\..\..\..\..\..\Data\Template.docx");

			// Insert image watermark
			InsertImageWatermark(document);

			// Save the modified document to a new file
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose the document object
			document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }

		private void InsertImageWatermark(Document document) {
			// Create a PictureWatermark object
			PictureWatermark picture = new PictureWatermark();
			// Load the image for the watermark
			picture.Picture = System.Drawing.Image.FromFile(@"..\..\..\..\..\..\Data\ImageWatermark.png");
			// Set the scaling of the watermark
			picture.Scaling = 250;
			// Specify whether the watermark should be washed out
			picture.IsWashout = false;
			// Set the watermark for the document
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
