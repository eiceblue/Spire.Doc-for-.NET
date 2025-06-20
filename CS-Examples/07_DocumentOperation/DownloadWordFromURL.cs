using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DownloadWordFromURL
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
			Document document = new Document();

			// Create a new instance of WebClient
			WebClient webClient = new WebClient();

			// Download the Word file from the specified URL and store it in a MemoryStream
			using (MemoryStream ms = new MemoryStream(webClient.DownloadData("http://www.e-iceblue.com/images/test.docx")))
			{
				// Load the document from the MemoryStream in Docx format
				document.LoadFromStream(ms, FileFormat.Docx);
			}

			// Specify the file name for the downloaded result
			String result = "Result-DownloadWordFileFromURL.docx";

			// Save the downloaded document to the specified file path in Docx2013 format
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose of the Document object to release resources
			document.Dispose();

            //Launch the MS Word file.
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
