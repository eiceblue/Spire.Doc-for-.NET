using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace HtmlToPdf
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_HtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

			string result = "Result-HtmlToPdf.pdf";

			//Save to file.
			document.SaveToFile(result, FileFormat.PDF);

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
