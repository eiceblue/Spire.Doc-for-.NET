using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ParagraphFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new empty document instance.
            Document doc = new Document();

            // Create a document navigator to help navigate and modify the document content.
            DocumentNavigator navigator = new DocumentNavigator(doc);

            // Load an existing Word document from the specified relative file path.
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            // Move the navigator's cursor to the first section of the document (section index 0).
            navigator.MoveToSection(0);

            // Move the cursor to the first paragraph (index 0) at character position 0 within that paragraph.
            navigator.MoveToParagraph(0, 0);

            // Set the line spacing rule for the current paragraph to "Multiple" (enables custom line spacing multiplier).
            navigator.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;

            // Set the line spacing to 1.5 times the default font size (assuming 12-point font: 1.5 * 12 = 18 points).
            navigator.ParagraphFormat.LineSpacing = 1.5F * 12F;

            // Set the left indent of the current paragraph to 5 points.
            navigator.ParagraphFormat.LeftIndent = 5;

            // Move the cursor to the third paragraph (index 2) at character position 0.
            navigator.MoveToParagraph(2, 0);

            // Set the background color of the current paragraph to blue.
            navigator.ParagraphFormat.BackColor = Color.Blue;

            // Save the modified document to a new file.
            doc.SaveToFile("ParagraphFormat.docx", FileFormat.Docx);

            // Close the document to release internal resources.
            doc.Close();

            // Explicitly dispose of the document object to free memory immediately.
            doc.Dispose();

            //Launch the MS Word file.
            WordDocViewer("ParagraphFormat.docx");
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
