using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveImageWatermark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
     
			// Create a new document object
			Document document = new Document();

			// Load the document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveImageWatermark.docx");

			// Remove the watermark from the document
			document.Watermark = null;

			// Specify the output file name
			String result = "Result-RemoveImageWatermark.docx";

			// Save the modified document to a new file in Docx2013 format
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose the document object
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
