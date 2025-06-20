using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace PreserveTheme
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Specify the input file path using a relative path.
			string input = @"..\..\..\..\..\..\..\Data\Theme.docx";

			// Create a new instance of the Document class.
			Document doc = new Document();

			// Load the document from the specified input file.
			doc.LoadFromFile(input);

			// Create another instance of the Document class.
			Document newWord = new Document();

			// Clone the default style, themes, and compatibility settings from the original document to the new document.
			doc.CloneDefaultStyleTo(newWord);
			doc.CloneThemesTo(newWord);
			doc.CloneCompatibilityTo(newWord);

			// Clone the first section from the original document and add it to the new document.
			newWord.Sections.Add(doc.Sections[0].Clone());

			// Specify the output file name.
			string output = "PreserveTheme.docx";

			// Save the new document to a file with the specified file name and file format as Docx.
			newWord.SaveToFile(output, FileFormat.Docx);

			// Dispose of the original document and the new document to release resources.
			doc.Dispose();
			newWord.Dispose();
			
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
