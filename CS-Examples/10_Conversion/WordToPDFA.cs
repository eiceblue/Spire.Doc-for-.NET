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
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            //Set the Conformance-level of the Pdf file to PDF_A1B.
            ToPdfParameterList toPdf = new ToPdfParameterList();
            toPdf.PdfConformanceLevel = Spire.Pdf.PdfConformanceLevel.Pdf_A1B;

            String result = "Result-WordToPDFA.pdf";

            //Save the file.
            document.SaveToFile(result, toPdf);

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
