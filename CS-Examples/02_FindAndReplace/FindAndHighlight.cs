using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace FindAndHighlight
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

            // Load an existing Word document from the specified file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

            // Find all occurrences of the string "word" in the document and retrieve their TextSelections
            TextSelection[] textSelections = document.FindAllString("word", false, true);

            // Iterate through each TextSelection
            foreach (TextSelection selection in textSelections)
            {
                // Get the entire range of the selection and set its CharacterFormat's HighlightColor property to Yellow
                selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
            }

            // Save the modified document to a file named "Sample.docx" in Docx format
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            // Dispose of the document object and release any associated resources
            document.Dispose();

            //Launching the  Word file.
            WordDocViewer("Sample.docx");
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
