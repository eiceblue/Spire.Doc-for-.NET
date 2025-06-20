using System;
using System.Windows.Forms;
using Spire.Doc;

namespace WordToMarkDown
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define the input file path relative to the current directory.
            string input = @"..\..\..\..\..\..\Data\ToMD.docx";

            // Create a new document
            Document doc = new Document();

            // Load .docx file
            doc.LoadFromFile(input);

            // Save to .md file
            doc.SaveToFile("WordToMarkDown_output.md", FileFormat.Markdown);

            // Dispose of the Document object
            doc.Close();

     
            WordDocViewer("WordToMarkDown_output.md");

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
