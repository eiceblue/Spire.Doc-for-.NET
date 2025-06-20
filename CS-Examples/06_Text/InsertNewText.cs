using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertNewText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Specify the input file path.
			string input = @"..\..\..\..\..\..\Data\Sample.docx";

			// Create a new instance of the Document class.
			Document doc = new Document();

			// Load a Word document from the specified input file path.
			doc.LoadFromFile(input);

			// Find all occurrences of the string "Word" within the document.
			// Perform a case-insensitive search and include whole word matches.
			TextSelection[] selections = doc.FindAllString("Word", true, true);

			// Initialize variables.
			int index = 0;
			TextRange range = new TextRange(doc);

			// Iterate through each found text selection.
			foreach (TextSelection selection in selections)
			{
				// Get the entire range of the selected text.
				range = selection.GetAsOneRange();

				// Create a new TextRange object with the document.
				TextRange newrange = new TextRange(doc);

				// Set the text of the new TextRange to "(New text)".
				newrange.Text = "(New text)";

				// Get the index of the range within its owner paragraph.
				index = range.OwnerParagraph.ChildObjects.IndexOf(range);

				// Insert the new TextRange after the current range in the owner paragraph.
				range.OwnerParagraph.ChildObjects.Insert(index + 1, newrange);
			}

			// Find all occurrences of the string "New text" within the document.
			// Perform a case-insensitive search and include whole word matches.
			TextSelection[] text = doc.FindAllString("New text", true, true);

			// Iterate through each found text selection.
			foreach (TextSelection selection in text)
			{
				// Set the highlight color of the text range to Yellow.
				selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
			}

			// Specify the file name for the resulting document.
			string output = "InsertNewText.docx";

			// Save the modified document to a file with the specified file name and format (Docx).
			doc.SaveToFile(output, FileFormat.Docx);

			// Clean up resources used by the document.
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
