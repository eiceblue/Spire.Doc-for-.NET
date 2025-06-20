using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
namespace AddDigitalSignature
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
			Document doc = new Document();

			// Load the document from the specified file path
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\AddDigitalSignature.doc");

			// Specify the output file path for the signed document with digital signature
			string result = "AddDigitalSignature_result.docx";

			// Save the document to the output file path in DOCX format with the specified certificate and password
			doc.SaveToFile(result, FileFormat.Docx, @"..\..\..\..\..\..\Data\gary.pfx", "e-iceblue");

			// Dispose the document object to free up resources
			doc.Dispose();

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
