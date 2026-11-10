using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace ToPdfWithDigitalSignature
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object to represent a Word document.
            Document document = new Document();

            // Load the Word document from the specified relative file path.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Sample.docx");

            // Create a parameter list object to configure PDF conversion settings.
            ToPdfParameterList toPdfParameter = new ToPdfParameterList();

            // Configure digital signature information using a PFX certificate file, password, reason, and location.
            toPdfParameter.DigitalSignatureInfo = new PdfDigitalSignatureInfo(@"..\..\..\..\..\..\..\Data\gary.pfx", "e-iceblue", "signing", "Spire Doc");

            // Set the hash algorithm used for the digital signature to RIPEMD-160.
            toPdfParameter.DigitalSignatureInfo.HashAlgorithm = PdfDigitalSignatureHashAlgorithm.RipeMD160;

            // Set the signature timestamp to the current date and time.
            toPdfParameter.DigitalSignatureInfo.SignatureDate = DateTime.Now;

            // Configure the timestamp server (TSA) settings for the digital signature, including URL, username, and password.
            toPdfParameter.DigitalSignatureInfo.TimestampConfiguration = new PdfDigitalSignatureTimestampConfiguration("https://freetsa.org/tsr", "Gary", "Password");

            // Save the document as a PDF file with the specified digital signature settings.
            document.SaveToFile("ToPdfWithDigitalSignature.pdf", toPdfParameter);

            // Close the document to release internal resources.
            document.Close();

            // Explicitly dispose of the document object to free memory immediately.
            document.Dispose();

            //view the PDF file.
            WordDocViewer("ToPdfWithDigitalSignature.pdf");
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
