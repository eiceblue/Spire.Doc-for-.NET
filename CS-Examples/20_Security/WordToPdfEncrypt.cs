using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace WordToPdfEncrypt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
      
			// Create a new Document object
			Document document = new Document();

			// Load the Word document file from the specified path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

			// Create a ToPdfParameterList object to specify PDF conversion parameters
			ToPdfParameterList toPdf = new ToPdfParameterList();

			// Encrypt the PDF with the specified password "e-iceblue"
			toPdf.PdfSecurity.Encrypt("e-iceblue");

			// Specify the output file name for the converted PDF
			String result = "Result-WordToPdfEncrypt.pdf";

			// Save the document as a PDF with the specified encryption settings
			document.SaveToFile(result, toPdf);

			// Dispose the Document object to free resources
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
