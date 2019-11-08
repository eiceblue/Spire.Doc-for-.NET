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
            //Open an html file.
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\InputHtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

            //Save it to a Word document.
            document.SaveToFile("HtmlFileToWord.docx", FileFormat.Docx);

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
