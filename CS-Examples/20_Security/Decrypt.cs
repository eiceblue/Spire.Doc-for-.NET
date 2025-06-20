using System;
using System.Windows.Forms;
using Spire.Doc;

namespace Decrypt
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

				// Load the document from the specified file path using the provided password
				document.LoadFromFile(@"..\..\..\..\..\..\Data\TemplateWithPassword.docx", FileFormat.Docx, "E-iceblue");

				// Save the document to the specified file path in DOCX format
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
