using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Formatting;
using Spire.Doc.Fields;

namespace SetParagraphTextDirection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new Document object.
            Document doc = new Document();

            // Add a new section to the document.
            Section section = doc.AddSection();

            // Add a new paragraph to the section.
            Paragraph paragraph = section.AddParagraph();

            // Append the text "Welcome to China." to the paragraph and get the TextRange object.
            TextRange farEastLayout = paragraph.AppendText("Welcome to China.");

            // Create a new FarEastLayout object to define vertical text settings.
            FarEastLayout style = new FarEastLayout();

            // Enable vertical text orientation for the layout style.
            style.Vertical = true;

            // Apply the vertical FarEastLayout style to the character format of the text range.
            farEastLayout.CharacterFormat.FarEastLayout = style;

            // Define the output file name for the saved document.
            String outputFile = "SetParagraphTextDirection.docx";

            // Save the document to the specified file in DOCX format.
            doc.SaveToFile(outputFile, FileFormat.Docx);

            // Close the document to release resources.
            doc.Close();

            WordDocViewer(outputFile);
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
