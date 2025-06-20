using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ChangeTOCStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
			// Create a new document
			Document doc = new Document();

			// Load the document from a file
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Toc.docx");

			// Create a custom Table of Contents (TOC) style
			ParagraphStyle tocStyle = Style.CreateBuiltinStyle(BuiltinStyle.Toc1, doc) as ParagraphStyle;
			tocStyle.CharacterFormat.FontName = "Aleo";
			tocStyle.CharacterFormat.FontSize = 15f;
			tocStyle.CharacterFormat.TextColor = Color.CadetBlue;
			doc.Styles.Add(tocStyle);

			// Iterate through all sections in the document
			foreach (Section section in doc.Sections)
			{
				// Iterate through all child objects in the body of each section
				foreach (DocumentObject obj in section.Body.ChildObjects)
				{
					// Check if the object is a StructureDocumentTag (e.g., TOC field)
					if (obj is StructureDocumentTag)
					{
						StructureDocumentTag tag = obj as StructureDocumentTag;
						
						// Iterate through all child objects within the StructureDocumentTag
						foreach (DocumentObject cObj in tag.ChildObjects)
						{
							// Check if the child object is a paragraph
							if (cObj is Paragraph)
							{
								Paragraph para = cObj as Paragraph;
								
								// Check if the paragraph has the style name "TOC1"
								if (para.StyleName == "TOC1")
								{
									// Apply the custom TOC style to the paragraph
									para.ApplyStyle(tocStyle.Name);
								}
							}
						}
					}
				}
			}

			// Specify the output file name
			string output = "ChangeTOCStyle_out.docx";

			// Save the modified document to a new file in DOCX format (version 2013)
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose of the document object
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
