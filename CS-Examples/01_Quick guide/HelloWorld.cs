using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace HelloWorld
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

            // Add a new Section to the document
            Section section = document.AddSection();

            // Add a new Paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            // Add text content to the paragraph
            paragraph.AppendText("Hello World!");

            // Save the document to a file named "Sample.docx" in Docx format
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            // Dispose of the document object and release any associated resources
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
