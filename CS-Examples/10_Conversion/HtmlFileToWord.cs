using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace HtmlFileToWord
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
			Document document = new Document();

			// Load an HTML file into the document object, with XHTML validation disabled.
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\InputHtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

			// Save the document as a DOCX file named "HtmlFileToWord.docx".
			document.SaveToFile("HtmlFileToWord.docx", FileFormat.Docx);

			// Dispose the document object to release resources.
			document.Dispose();

            //Launch the file.
            WordDocViewer("HtmlFileToWord.docx");
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
