using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace ToEpub
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           // Create a new instance of the Document class.
			Document doc = new Document();

			// Load a document from the specified file path.
			doc.LoadFromFile(@"..\..\..\..\..\..\..\Data\ToEpub.doc");

			// Specify the output file name for the EPUB file.
			string result = "result.epub";

			// Save the document as an EPUB file with the specified output file name and format.
			doc.SaveToFile(result, FileFormat.EPub);

			// Dispose the document object to release resources.
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
