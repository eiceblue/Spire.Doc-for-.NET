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
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

            //Create an instance of ToPdfParameterList.
            ToPdfParameterList toPdf = new ToPdfParameterList();

            //Set the user password for the resulted PDF file.
            toPdf.PdfSecurity.Encrypt("e-iceblue");          

            String result = "Result-WordToPdfEncrypt.pdf";

            //Save to file.
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
