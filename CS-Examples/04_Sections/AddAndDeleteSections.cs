using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddAndDeleteSections
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
            Document doc = new Document();

            // Load the document from the specified file path.
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\SectionTemplate.docx");

            // Call the AddSection method to add a new section to the document.
            AddSection(doc);

            // Call the DeleteSection method to remove the last section from the document.
            DeleteSection(doc);

            // Specify the output file name.
            string output = "AddAndDeleteSections_out.docx";

            // Save the modified document to the specified file in DOCX 2013 format.
            doc.SaveToFile(output, FileFormat.Docx2013);

            // Dispose of the document object to release resources.
            doc.Dispose();

            // Open the output file using a file viewer.
            FileViewer(output);
        }

        // Method to add a section to the document.
        private void AddSection(Document doc)
        {
            doc.AddSection();
        }

        // Method to delete the last section from the document.
        private void DeleteSection(Document doc)
        {
            doc.Sections.RemoveAt(doc.Sections.Count - 1);
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
