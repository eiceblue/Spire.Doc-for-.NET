using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace WordToPDFA
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

			//Create a ToPdfParameterList
			ToPdfParameterList toPdf = new ToPdfParameterList();

			//Set the Conformance-level of the Pdf file to PDF_A1B.
			toPdf.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;

			string result = "Result-WordToPDFA.pdf";

			//Save the file.
			document.SaveToFile(result, toPdf);

			//Dispose the document
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
