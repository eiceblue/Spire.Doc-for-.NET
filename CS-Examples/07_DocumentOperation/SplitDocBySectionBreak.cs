using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SplitDocBySectionBreak
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class
			Document document = new Document();

			// Load a Word document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_4.docx");

			// Declare a new Document object
			Document newWord;

			// Iterate through each section in the document
			for (int i = 0; i < document.Sections.Count; i++)
			{
				// Specify the file name for the result document using the section index
				String result = String.Format("Result-SplitWordFileBySectionBreak_{0}.docx", i);

				// Create a new instance of the Document class to hold the split section
				newWord = new Document();

				// Clone the section at the current index and add it to the new document
				newWord.Sections.Add(document.Sections[i].Clone());

				// Save the new document with the split section to a file
				newWord.SaveToFile(result);

				// Dispose of the original and new document objects to release resources
				document.Dispose();
				newWord.Dispose();

                //Launch the MS Word file.
                WordDocViewer(result);
            }
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
