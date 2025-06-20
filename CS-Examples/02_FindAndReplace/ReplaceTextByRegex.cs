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
            // Create a new instance of the Document class
            Document doc = new Document();

            // Load a Word document from a file
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextByRegex.docx");

            // Create a regular expression object to search for a pattern
            Regex regex = new Regex(@"\#\w+\b");

            // Replace all occurrences of the pattern in the document with the string "Spire.Doc"
            doc.Replace(regex, "Spire.Doc");

            // Save the modified document to a file
            doc.SaveToFile("output.docx", FileFormat.Docx);

            // Dispose the document and release all resources it is using
            doc.Dispose();

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
