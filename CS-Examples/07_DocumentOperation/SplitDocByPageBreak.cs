using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SplitDocByPageBreak
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class to hold the original document
			Document original = new Document();

			// Load the Word document from the specified file path
			original.LoadFromFile(@"..\..\..\..\..\..\..\Data\SplitWordFileByPageBreak.docx");

			// Create a new instance of the Document class to hold the modified document
			Document newWord = new Document();

			// Add a section to the new document
			Section section = newWord.AddSection();

			// Clone the default style, themes, and compatibility settings from the original document to the new document
			original.CloneDefaultStyleTo(newWord);
			original.CloneThemesTo(newWord);
			original.CloneCompatibilityTo(newWord);

			// Initialize an index variable to keep track of the split documents
			int index = 0;

			// Iterate through each section in the original document
			foreach (Section sec in original.Sections)
			{
				// Iterate through each object in the body of the section
				foreach (DocumentObject obj in sec.Body.ChildObjects)
				{
					// Check if the object is a paragraph
					if (obj is Paragraph)
					{
						// Cast the object as a Paragraph
						Paragraph para = obj as Paragraph;

						// Clone the section properties from the original section to the new section
						sec.CloneSectionPropertiesTo(section);

						// Add the cloned paragraph to the body of the new section
						section.Body.ChildObjects.Add(para.Clone());

						// Iterate through each object in the child objects of the paragraph
						foreach (DocumentObject parobj in para.ChildObjects)
						{
							// Check if the object is a page break
							if (parobj is Break && (parobj as Break).BreakType == BreakType.PageBreak)
							{
								// Get the index of the page break within the paragraph
								int i = para.ChildObjects.IndexOf(parobj);

								// Remove the page break from the last paragraph in the section
								section.Body.LastParagraph.ChildObjects.RemoveAt(i);

								// Save the split document to a file with an incremented index
								newWord.SaveToFile(String.Format("Result-SplitWordFileByPageBreak-{0}.docx", index), FileFormat.Docx);

								// Increment the index for the next split document
								index++;

								// Create a new instance of the Document class for the next split document
								newWord = new Document();

								// Add a section to the new document
								section = newWord.AddSection();

								// Clone the default style, themes, and compatibility settings from the original document to the new document
								original.CloneDefaultStyleTo(newWord);
								original.CloneThemesTo(newWord);
								original.CloneCompatibilityTo(newWord);

								// Clone the section properties from the original section to the new section
								sec.CloneSectionPropertiesTo(section);

								// Add the cloned paragraph to the body of the new section
								section.Body.ChildObjects.Add(para.Clone());

								// Check if the first paragraph in the section is empty and remove it if necessary
								if (section.Paragraphs[0].ChildObjects.Count == 0)
								{
									section.Body.ChildObjects.RemoveAt(0);
								}
								else
								{
									// Remove all objects before the page break in the first paragraph of the section
									while (i >= 0)
									{
										section.Paragraphs[0].ChildObjects.RemoveAt(i);
										i--;
									}
								}
							}
						}
					}
					
					// Check if the object is a table and add it to the body of the section
					if (obj is Table)
					{
						section.Body.ChildObjects.Add(obj.Clone());
					}
				}
			}

			// Specify the file name for the result document
			string result = string.Format("Result-SplitWordFileByPageBreak-{0}.docx", index);

			// Save the final modified document to the specified file path in the Docx2013 format
			newWord.SaveToFile(result, FileFormat.Docx2013);

			// Dispose of the original and new document objects to release resources
			original.Dispose();
			newWord.Dispose();

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
