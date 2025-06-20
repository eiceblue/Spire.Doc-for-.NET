using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace KeepHiddenText
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

			// Load a Word document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_5.docx");

			// Create a ToPdfParameterList object to specify conversion parameters
			ToPdfParameterList pdf = new ToPdfParameterList();

			// Set the 'IsHidden' parameter to true, which hides any hidden text in the converted PDF
			pdf.IsHidden = true;

			// Specify the output file name for the converted PDF
			string result = "Result-SaveTheHiddenTextToPDF.pdf";

			// Save the document as a PDF using the specified conversion parameters
			document.SaveToFile(result, pdf);

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
