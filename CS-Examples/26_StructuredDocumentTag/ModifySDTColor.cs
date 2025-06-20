using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace ModifySDTColor
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

			// Load a document from the specified file path
			doc.LoadFromFile(@"..\..\..\..\..\..\Data\ModifySTDColor.docx");

			// Iterate through the sections in the document
			for (int s = 0; s < doc.Sections.Count; s++)
			{
				// Get the current section
				Section section = doc.Sections[s];

			// Iterate through the child objects in the section's body
			for (int i = 0; i < section.Body.ChildObjects.Count; i++)
			{
				// Check if the child object is a Paragraph
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
							
							// Get the SDTProperties of the StructureDocumentTagInline
							SDTProperties sDTProperties = sdt.SDTProperties;

							// Set the color of the SDTProperties based on the SDTType
							switch (sDTProperties.SDTType)
							{
								case SdtType.RichText:
									sDTProperties.Color = Color.Orange;
									break;
								case SdtType.Text:
									sDTProperties.Color = Color.Green;
									break;
							}
						}
					}
				}

				// Check if the child object is a StructureDocumentTag
				if (section.Body.ChildObjects[i] is StructureDocumentTag)
				{
					// Get the StructureDocumentTag object
					StructureDocumentTag sdt = section.Body.ChildObjects[i] as StructureDocumentTag;
					
					// Get the SDTProperties of the StructureDocumentTag
					SDTProperties sDTProperties = sdt.SDTProperties;

					// Set the color of the SDTProperties based on the SDTType
					switch (sDTProperties.SDTType)
					{
						case SdtType.RichText:
							sDTProperties.Color = Color.Orange;
							break;
						case SdtType.Text:
							sDTProperties.Color = Color.Green;
							break;
					}
				}
			}
			}

			// Specify the output file path
			string output = "ModifySTDColor_out.docx";

			// Save the modified document to the output file in DOCX format
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
