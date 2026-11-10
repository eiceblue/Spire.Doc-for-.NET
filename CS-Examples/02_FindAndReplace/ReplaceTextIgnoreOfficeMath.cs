using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields.Shapes.Charts;
using Spire.Doc.Interface;

namespace ReplaceTextIgnoreOfficeMath
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new Document object to work with Word documents.
            Document document = new Document();

            // Load an existing Word document from the specified file path into the document object.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextIgnoreOfficeMath.docx");

            // Create a new SearchOptions instance to customize search behavior during text replacement.
            SearchOptions searchOptions = new SearchOptions();

            // Configure the search to skip Office Math objects (e.g., equations) during replacement.
            searchOptions.IgnoreOfficeMath = true;

            // Ensure that only whole words matching the search term are replaced (not partial matches).
            searchOptions.FindWholeWordsOnly = true;

            // Replace all occurrences of the word "a" with "an" in the document, using the defined search options.
            document.Replace("a", "an", searchOptions);

            // Save the modified document to a file in Word 2016 format.
            document.SaveToFile("ReplaceTextIgnoreOfficeMath_result.docx", FileFormat.Docx2016);

            // Close the document to release internal resources such as file handles and memory buffers.
            document.Close();

            // Explicitly dispose of the document object to immediately free unmanaged resources and memory.
            document.Dispose();


            //Launching the Word file.
            WordDocViewer("ReplaceTextIgnoreOfficeMath_result.docx");
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
