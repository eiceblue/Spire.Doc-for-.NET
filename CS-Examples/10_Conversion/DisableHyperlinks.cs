using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DisableHyperlinks
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

			// Load a Word document from a specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_5.docx");

			// Create a ToPdfParameterList object to specify conversion parameters for PDF export
			ToPdfParameterList pdf = new ToPdfParameterList();
			
			//Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
			pdf.DisableLink = true;

			// Specify the output file name for the converted PDF
			string result = "Result-DisableHyperlinks.pdf";

			// Save the document to PDF format with the specified parameters and file name
			document.SaveToFile(result, pdf);

			// Dispose of the Document object to release resources
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
