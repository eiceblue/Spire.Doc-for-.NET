using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SpecifiedProtectionType
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

			// Load the Word document file from the specified path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

			// Set the document protection to allow only reading with the specified password
			document.Protect(ProtectionType.AllowOnlyReading, "123456");

			// Specify the output file name for the modified document
			String result = "Result-SpecifiedProtectionType.docx";

			// Save the modified document to the specified file format
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose the Document object to free resources
			document.Dispose();

            //Launch the file.
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
