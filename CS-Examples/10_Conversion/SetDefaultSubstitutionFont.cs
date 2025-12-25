using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetDefaultSubstitutionFont
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

            // Set the default substitution font name to "Arial"
            // This font will be used if a specified font is not available
            document.DefaultSubstitutionFontName = "Arial";

            // Add a new section to the document
            Section section = document.AddSection();

            // Add a new paragraph to the section
            Paragraph paragraph = section.AddParagraph();

            // Append text to the paragraph and get a reference to the text range
            TextRange textRange = paragraph.AppendText("Welcome to evaluate Spire.Doc for .NET product.");

            // Set the font name of the text range to "San Francisco"
            // (This font might not be available on the system)
            textRange.CharacterFormat.FontName = "San Francisco";

            // Set the font size of the text range to 16
            textRange.CharacterFormat.FontSize = 16;

            // Define the output file name
            string result = "SetDefaultSubstitutionFont-result.pdf";

            // Save the document to a PDF file
            document.SaveToFile(result, FileFormat.PDF);

            // Dispose of the Document object to release resources
            document.Dispose();

            //Launching the MS Word file.
            WordDocViewer(result);
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
