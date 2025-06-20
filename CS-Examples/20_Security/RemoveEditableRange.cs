using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace RemoveEditableRange
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

			// Load the document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveEditableRange.docx");

			// Iterate through each section in the document
			foreach (Section section in document.Sections)
			{
				// Iterate through each paragraph in the section's body
				foreach (Paragraph paragraph in section.Body.Paragraphs)
				{
					// Loop through the child objects of the paragraph
					for (int i = 0; i < paragraph.ChildObjects.Count; )
					{
						DocumentObject obj = paragraph.ChildObjects[i];
						
						// Check if the child object is a PermissionStart or PermissionEnd element
						if (obj is PermissionStart || obj is PermissionEnd)
						{
							// Remove the PermissionStart or PermissionEnd element from the paragraph
							paragraph.ChildObjects.Remove(obj);
						}
						else
						{
							// Move to the next child object
							i++;
						}
					}
				}
			}

			// Specify the output file path for the modified document
			string output = "RemoveEditableRange_output.docx";

			// Save the modified document to the output file path in DOCX format
			document.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to free up resources
			document.Dispose();
			
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
