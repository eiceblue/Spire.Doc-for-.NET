using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Text;

namespace AddBarcodeImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load a document from disk
			Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_2.docx");

			//Specidied the image path
			string imgPath = @"..\..\..\..\..\..\Data\barcode.png";

			//Add barcode image
			DocPicture picture = document.Sections[0].AddParagraph().AppendPicture(Image.FromFile(imgPath));

			//Save to file
			document.SaveToFile("result.docx", FileFormat.Docx);

			//Dispose the document
			document.Dispose();

            //Launch result file
            WordDocViewer("result.docx");

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
