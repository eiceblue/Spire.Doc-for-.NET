using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetImageQuality
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
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Doc_1.doc", FileFormat.Doc);

            //Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
            document.JPEGQuality = 40;

            String result = "Result-DocToPDFImageQuality.pdf";

            //Save to file.
            document.SaveToFile(result, FileFormat.PDF);

            //Launch the Pdf file.
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
