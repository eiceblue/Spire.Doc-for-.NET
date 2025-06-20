using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Text.RegularExpressions;

namespace RemoveTableOfContent
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

			// Load a Word document from a specific file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\TableOfContent.docx");

			// Access the body of the first section in the document
			Body body = document.Sections[0].Body;

			// Define a regular expression pattern to match the style names
			Regex regex = new Regex("TOC\\w+");

			// Iterate over the paragraphs in the body
			for (int i = 0; i < body.Paragraphs.Count; i++)
			{
				// Check if the style name matches the regular expression pattern
				if (regex.IsMatch(body.Paragraphs[i].StyleName))
				{
					// Remove the paragraph if it matches the pattern
					body.Paragraphs.RemoveAt(i);
					
					// Decrement the counter to avoid skipping the next paragraph
					i--;
				}
			}

			// Save the modified document to a new file named "Output.docx"
			document.SaveToFile("Output.docx", FileFormat.Docx);

			// Dispose the document object to free up resources
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
