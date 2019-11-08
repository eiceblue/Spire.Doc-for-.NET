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
            //Get html string.
            String HTML = File.ReadAllText(@"..\..\..\..\..\..\..\Data\InputHtml.txt");

			//Create a new document.
            Document document = new Document();

            //Add a section.
            Section sec = document.AddSection();

            //Add a paragraph and append html string.
            sec.AddParagraph().AppendHTML(HTML);

            //Save it to a Word file.
            document.SaveToFile("HtmlFileToWord.docx", FileFormat.Docx);

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
