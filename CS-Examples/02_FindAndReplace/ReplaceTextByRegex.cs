using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.Text.RegularExpressions;

namespace ReplaceTextByRegex
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //create a document
            Document doc = new Document();

            //Load the document from disk.
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextByRegex.docx");

            //create a regex, match the text that starts with #
            Regex regex = new Regex(@"\#\w+\b");

            //replace the text by regex
            doc.Replace(regex, "Spire.Doc");

            //save the document
            doc.SaveToFile("output.docx", FileFormat.Docx);

            //view the document
            WordDocViewer("output.docx");

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
