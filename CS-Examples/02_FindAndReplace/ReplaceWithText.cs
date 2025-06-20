using System;
using System.Windows.Forms;
using Spire.Doc;

namespace Replace
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
            Document document = new Document();

            // Load the content of the document from the given path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            // Replace all occurrences of the word "word" with the text "ReplacedText"
            document.Replace("word", "ReplacedText", false, true);

            // Save the modified document to the given filename with the .docx file format
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            // Dispose of the Document object to release its resources
            document.Dispose();

            //Launching the MS Word file.
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
