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
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_5.docx");

            //Create an instance of ToPdfParameterList.
            ToPdfParameterList pdf = new ToPdfParameterList();

            //Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
            //Set DisableLink to false to preserve the hyperlink effect for the result PDF page.
            pdf.DisableLink = true;

            String result = "Result-DisableHyperlinks.pdf";

            //Save to file.
            document.SaveToFile(result, pdf);

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
