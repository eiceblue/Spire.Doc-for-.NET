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
            // Create a new instance of the Document class
			Document document = new Document();

			// Load a Word document from the specified file path using the 'Doc' format
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Doc_1.doc", FileFormat.Doc);

			// Set the JPEG quality for saving images in the document to 40%
			document.JPEGQuality = 40;

			// Specify the output file name for the converted PDF
			string result = "Result-DocToPDFImageQuality.pdf";

			// Save the document as a PDF using the specified file format
			document.SaveToFile(result, FileFormat.PDF);

			// Dispose of the document object to free up resources
			document.Dispose();

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
