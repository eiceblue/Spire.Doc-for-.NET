using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CloneSection
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
            Document srcDoc = new Document();

            // Load a Word document from a specified file path using the LoadFromFile method.
            srcDoc.LoadFromFile(@"..\..\..\..\..\..\Data\SectionTemplate.docx");

            // Create a new instance of the Document class.
            Document desDoc = new Document();

            // Initialize a variable to store a cloned section.
            Section cloneSection = null;

            // Iterate through each section in the source document.
            foreach (Section section in srcDoc.Sections)
            {
                // Clone the current section and assign it to the cloneSection variable.
                cloneSection = section.Clone();

                // Add the cloned section to the destination document's Sections collection.
                desDoc.Sections.Add(cloneSection);
            }

            // Specify the output file name for the cloned section document.
            string output = "CloneSection_out.docx";

            // Save the destination document to a file with the specified output file name and Docx2013 format.
            desDoc.SaveToFile(output, FileFormat.Docx2013);

            // Release system resources used by the source and destination documents.
            srcDoc.Dispose();
            desDoc.Dispose();

            //Launch Word file
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