using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ChangeCase
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
			string input = @"..\..\..\..\..\..\Data\Text1.docx";

			// Create a new instance of the Document class.
			Document doc = new Document();

			// Load a Word document from the specified input file path.
			doc.LoadFromFile(input);

			TextRange textRange;

			// Get the second paragraph in the first section of the document.
			Paragraph para1 = doc.Sections[0].Paragraphs[1];

			// Iterate through each child object within the paragraph.
			foreach (DocumentObject obj in para1.ChildObjects)
			{
				// Check if the child object is a TextRange.
				if (obj is TextRange)
				{
					// Cast the child object to a TextRange and set the AllCaps property to true,
					// which converts the text to all capital letters.
					textRange = obj as TextRange;
					textRange.CharacterFormat.AllCaps = true;
				}
			}

			// Get the fourth paragraph in the first section of the document.
			Paragraph para2 = doc.Sections[0].Paragraphs[3];

			// Iterate through each child object within the paragraph.
			foreach (DocumentObject obj in para2.ChildObjects)
			{
				// Check if the child object is a TextRange.
				if (obj is TextRange)
				{
					// Cast the child object to a TextRange and set the IsSmallCaps property to true,
					// which converts the text to small capital letters.
					textRange = obj as TextRange;
					textRange.CharacterFormat.IsSmallCaps = true;
				}
			}

			// Specify the file name for the resulting document.
			string output = "ChangeCase.docx";

			// Save the modified document to a file with the specified file name and format (Docx2013).
			doc.SaveToFile(output, FileFormat.Docx2013);

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
