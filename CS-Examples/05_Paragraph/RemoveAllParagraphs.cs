using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveAllParagraphs
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            // Iterate through each section in the document.
            foreach (Section section in document.Sections)
            {
                // Clear all paragraphs within the section.
                section.Paragraphs.Clear();
            }

            // Specify the file name for the resulting document.
            String result = "Result-RemoveAllParagraphs.docx";

            // Save the modified document to a file with the specified file name and format (Docx2013).
            document.SaveToFile(result, FileFormat.Docx2013);

            // Clean up resources used by the document.
            document.Dispose();

            //Launch the MS Word file.
            WordDocViewer(result);
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
