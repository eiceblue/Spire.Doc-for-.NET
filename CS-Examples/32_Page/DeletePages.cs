using System;
using System.Windows.Forms;
using Spire.Doc;

namespace DeletePages
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize a new Document object
            Document document = new Document();

            // Load an existing Word document from the specified relative file path
            document.LoadFromFile(@"..\..\..\..\..\..\Data\RemovePages.docx");

            // Remove all blank pages from the document
            document.RemoveBlankPages();

            // Remove specific pages by index (0-based). Here, it removes the 3rd page (index 2) and the 5th page (index 4).
            document.RemovePages(new System.Collections.Generic.List<int> { 2, 4 });

            // Define the output file name for the modified document
            String outputFile = "DeletePages.docx";

            // Save the document to the specified file in DOCX 2019 format
            document.SaveToFile(outputFile, FileFormat.Docx2019);

            // Close the document to release file handles
            document.Close();

            // Dispose of the document object to free up memory
            document.Dispose();

            //Launch the Word file.
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
