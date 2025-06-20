using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.IO;

namespace SimpleInsertFile
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

			// Load the document from the specified file path.
			doc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N5.docx");

			// Insert text from another document into the current document, specifying the file path and allowing automatic file format detection.
			doc.InsertTextFromFile(@"..\..\..\..\..\..\..\Data\Template_N3.docx", FileFormat.Auto);

			// Specify the output file name.
			string output = "SimpleInsertFile_out.docx";

			// Save the modified document to a file with the specified file name and file format as Docx2013.
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose of the document to release resources.
			doc.Dispose();

            //Launch the document
            WordDocViewer(output);
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
