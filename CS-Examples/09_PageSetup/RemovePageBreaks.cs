using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemovePageBreaks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new Document object.
			Document document = new Document();

			// Load an existing document from a file.
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_4.docx");

			// Iterate through paragraphs in the first section of the document.
			for (int j = 0; j < document.Sections[0].Paragraphs.Count; j++)
			{
				// Get a reference to the current paragraph.
				Paragraph p = document.Sections[0].Paragraphs[j];

				// Iterate through child objects (elements) within the paragraph.
				for (int i = 0; i < p.ChildObjects.Count; i++)
				{
					// Get a reference to the current child object.
					DocumentObject obj = p.ChildObjects[i];

					// Check if the child object is a Break.
					if (obj.DocumentObjectType == DocumentObjectType.Break)
					{
						// Remove the Break from the paragraph's child objects.
						Break b = obj as Break;
						p.ChildObjects.Remove(b);
					}
				}
			}

			// Specify the filename for the resulting document without page breaks.
			string result = "Result-RemovePageBreaks.docx";

			// Save the modified document to a file in the Docx2013 format.
			document.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources associated with the document.
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
