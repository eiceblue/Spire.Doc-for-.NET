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

namespace RemoveFootnote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			// Create a new instance of Document
			Document document = new Document();

			// Load the Word document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Footnote.docx");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Iterate through each paragraph in the section
			foreach (Paragraph para in section.Paragraphs)
			{
				int index = -1;
				
				// Find the index of the first footnote within the paragraph's child objects
				for (int i = 0, cnt = para.ChildObjects.Count; i < cnt; i++)
				{
					ParagraphBase pBase = para.ChildObjects[i] as ParagraphBase;
					
					if (pBase is Footnote)
					{
						index = i;
						break;
					}
				}

				// If a footnote is found, remove it from the paragraph's child objects
				if (index > -1)
					para.ChildObjects.RemoveAt(index);
			}

			// Save the modified document to a file
			document.SaveToFile("RemoveFootnote.docx", FileFormat.Docx);

			// Dispose of the document object when finished using it
			document.Dispose();

            //view the Word file.
            WordDocViewer("RemoveFootnote.docx");
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
