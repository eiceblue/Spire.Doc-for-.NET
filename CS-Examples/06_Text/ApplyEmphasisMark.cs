using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ApplyEmphasisMark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           // Create a new instance of the Document class.
			Document document = new Document();

			// Load a Word document from a specified file path.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			// Find all occurrences of the string "Spire.Doc for .NET" in the document.
			TextSelection[] textSelections = document.FindAllString("Spire.Doc for .NET", false, true);

			// Iterate through each text selection result.
			foreach (TextSelection selection in textSelections)
			{
				// Get the found text range as a single range and apply an emphasis mark (dot) to its character format.
				selection.GetAsOneRange().CharacterFormat.EmphasisMark = Emphasis.Dot;
			}

			// Specify the file name for the resulting document.
			string output = "ApplyEmphasisMark.docx";

			// Save the modified document to a file with the specified file name and format (Docx).
			document.SaveToFile(output, FileFormat.Docx);

			// Clean up resources used by the document.
			document.Dispose();

            //Launching the file
            WordDocViewer(output);
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
