using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RtfToHtml
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
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_RtfFile.rtf");

			string result = "Result-RtfToHtml.html";

			//Save to file.
			document.SaveToFile(result, FileFormat.Html);

			//Dispose the document
			document.Dispose();

            //Launch the Html file.
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
