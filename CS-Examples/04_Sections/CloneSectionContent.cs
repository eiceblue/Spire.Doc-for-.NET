using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CloneSectionContent
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
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\SectionTemplate.docx");

            // Get the first section from the document and assign it to the sec1 variable.
            Section sec1 = doc.Sections[0];

            // Get the second section from the document and assign it to the sec2 variable.
            Section sec2 = doc.Sections[1];

            // Iterate through each DocumentObject in the child objects collection of sec1's Body.
            foreach (DocumentObject obj in sec1.Body.ChildObjects)
            {
                // Clone the current DocumentObject and add it to the child objects collection of sec2's Body.
                sec2.Body.ChildObjects.Add(obj.Clone());
            }

            // Specify the output file name for the cloned section content document.
            string output = "CloneSectionContent_out.docx";

            // Save the document to a file with the specified output file name and Docx2013 format.
            doc.SaveToFile(output, FileFormat.Docx2013);

            // Release system resources used by the document.
            doc.Dispose();

            //Launch the file
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
