using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace SetUnderlineColor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document instance
            Document document = new Document();

            // Add a new section to the document
            Section section = document.AddSection();

            // Add a new paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            // Append text to the paragraph and get the TextRange object for formatting
            TextRange textRange = paragraph.AppendText("Welcome to evaluate Spire.Doc for .NET product.");

            // Set the underline style of the text to single underline
            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            // Set the underline color of the text to red
            textRange.CharacterFormat.UnderlineColor = Color.Red;

            // Define the file path and name for saving the document
            string filePath = "SetUnderlineColor.docx";

            // Save the document to the specified file path in DOCX format
            document.SaveToFile(filePath, FileFormat.Docx);

            // Release resources used by the document
            document.Dispose();

            WordDocViewer(filePath);
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
