using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Encrypt
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

			// Load the document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template.docx");

			// Encrypt the document with the provided password
			document.Encrypt("E-iceblue");

			// Save the encrypted document to the specified file path in DOCX format
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose the document object to free up resources
			document.Dispose();

            //Launching the Word file.
            WordDocViewer("Sample.docx");


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
