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
            //Create Word document.
            Document document = new Document();

            //Initialize a new instance of WebClient class.
            WebClient webClient = new WebClient();

            //Download a Word document from URL.
            using (MemoryStream ms = new MemoryStream(webClient.DownloadData("http://www.e-iceblue.com/images/test.docx")))
            {
                document.LoadFromStream(ms, FileFormat.Docx);
            }

            String result = "Result-DownloadWordFileFromURL.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

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
