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

namespace RemoveSectionContent
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

            // Load a Word document from a specified file path using the LoadFromFile method.
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_N3.docx");

            // Iterate through each section in the document.
            foreach (Section section in doc.Sections)
            {
                // Clear the child objects in the header of the current section.
                section.HeadersFooters.Header.ChildObjects.Clear();

                // Clear the child objects in the body of the current section.
                section.Body.ChildObjects.Clear();

                // Clear the child objects in the footer of the current section.
                section.HeadersFooters.Footer.ChildObjects.Clear();
            }

            // Specify the output file name for the section content removal document.
            string output = "RemoveSectionContent_out.docx";

            // Save the document to a file with the specified output file name and Docx2013 format.
            doc.SaveToFile(output, FileFormat.Docx2013);

            // Release system resources used by the document.
            doc.Dispose();

            //Launch the file
            FileViewer(output);
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
