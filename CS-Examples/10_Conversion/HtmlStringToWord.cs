using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.IO;

namespace HtmlStringToWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Read the content of an HTML file into a string variable.
			string HTML = File.ReadAllText(@"..\..\..\..\..\..\..\Data\InputHtml.txt");

			// Create a new instance of the Document class.
			Document document = new Document();

			// Add a new section to the document.
			Section sec = document.AddSection();

			// Add a new paragraph to the section and append the HTML content to it.
			sec.AddParagraph().AppendHTML(HTML);

			// Save the document as a DOCX file named "HtmlFileToWord.docx".
			document.SaveToFile("HtmlFileToWord.docx", FileFormat.Docx);

			// Dispose the document object to release resources.
			document.Dispose();

            //Launch the Word file.
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
