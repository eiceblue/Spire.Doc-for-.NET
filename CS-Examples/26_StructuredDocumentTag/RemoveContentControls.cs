using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace RemoveContentControls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
			// Create a new document object
			Document doc = new Document();

			// Load a document file from a specified path
			doc.LoadFromFile(@"...\..\..\..\..\..\Data\RemoveContentControls.docx");

			// Iterate through the sections in the document
			for (int s = 0; s < doc.Sections.Count; s++)
			{
				// Get the current section
				Section section = doc.Sections[s];

			// Iterate through the child objects in the section's body
			for (int i = 0; i < section.Body.ChildObjects.Count; i++)
			{
				// Check if the child object is a paragraph
				if (section.Body.ChildObjects[i] is Paragraph)
				{
					// Get the paragraph object
					Paragraph para = section.Body.ChildObjects[i] as Paragraph;
					
					// Iterate through the child objects in the paragraph
					for (int j = 0; j < para.ChildObjects.Count; j++)
					{
						// Check if the child object is a StructureDocumentTagInline
						if (para.ChildObjects[j] is StructureDocumentTagInline)
						{
							// Get the StructureDocumentTagInline object
							StructureDocumentTagInline sdt = para.ChildObjects[j] as StructureDocumentTagInline;
							
							// Remove the StructureDocumentTagInline from the paragraph
							para.ChildObjects.Remove(sdt);
							
							// Decrement the index to account for the removed object
							j--;
						}
					}
				}
				
				// Check if the child object is a StructureDocumentTag
				if (section.Body.ChildObjects[i] is StructureDocumentTag)
				{
					// Get the StructureDocumentTag object
					StructureDocumentTag sdt = section.Body.ChildObjects[i] as StructureDocumentTag;
					
					// Remove the StructureDocumentTag from the section's body
					section.Body.ChildObjects.Remove(sdt);
					
					// Decrement the index to account for the removed object
					i--;
				}
			}
			}

			// Save the modified document to a new file
			string output = "RemoveContentControls_out.docx";
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose the document object
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
