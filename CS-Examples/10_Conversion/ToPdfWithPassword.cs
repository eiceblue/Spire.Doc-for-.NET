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
            // Create a new Document instance
			Document document = new Document();

			// Load the Word document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ConvertedTemplate.docx");

			// Create a ToPdfParameterList instance to configure PDF conversion options
			ToPdfParameterList toPdf = new ToPdfParameterList();

			// Set a password for the PDF encryption
			string password = "E-iceblue";
			toPdf.PdfSecurity.Encrypt(password, password, PdfPermissionsFlags.Default, PdfEncryptionKeySize.Key128Bit);

			// Save the document as a PDF file with encryption and the specified output file name
			document.SaveToFile("EncryptWithPassword.pdf", toPdf);

			// Dispose the Document object after use
			document.Dispose();

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
