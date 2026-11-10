using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ListFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new document instance.
            Document doc = new Document();

            // Add a new section to the document.
            Section section = doc.AddSection();

            // Add the first paragraph to the section.
            Paragraph paragraph = section.AddParagraph();

            // Append text to the first paragraph.
            paragraph.AppendText("This is the first paragraph.");

            // Add a second paragraph to the section.
            paragraph = section.AddParagraph();

            // Append text to the second paragraph.
            paragraph.AppendText("This is the second paragraph.");

            // Add a third paragraph to the section.
            paragraph = section.AddParagraph();

            // Append text to the third paragraph.
            paragraph.AppendText("This is the third paragraph.");

            // Create a document navigator to facilitate navigation and formatting.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Apply bullet list style to the current position (first paragraph by default).
            navigator.ListFormat.ApplyBulletStyle();

            // Move the navigator's cursor to the third paragraph (index 2) at character offset 0.
            navigator.MoveToParagraph(2, 0);

            // Apply bullet list style to the third paragraph.
            navigator.ListFormat.ApplyBulletStyle();

            // Save the document to a file named "ListFormat.docx" in DOCX format.
            doc.SaveToFile("ListFormat.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("ListFormat.docx");
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
