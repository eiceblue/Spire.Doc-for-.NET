using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace ToPdfWithPassword
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //create word document
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ConvertedTemplate.docx");

            //create a parameter
            ToPdfParameterList toPdf = new ToPdfParameterList();

            //set the password
            string password = "E-iceblue";
            toPdf.PdfSecurity.Encrypt(password, password, Spire.Pdf.Security.PdfPermissionsFlags.Default, Spire.Pdf.Security.PdfEncryptionKeySize.Key128Bit);        
            //save doc file.
            document.SaveToFile("EncryptWithPassword.pdf", toPdf);

            //view the PDF file.
            WordDocViewer("EncryptWithPassword.pdf");
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
