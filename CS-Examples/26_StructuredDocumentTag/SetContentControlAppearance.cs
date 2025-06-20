using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetContentControlAppearance
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
 
			// Specify the input file path
			string input = @"..\..\..\..\..\..\Data\ContentControl.docx";

			// Create a new document object
			Document doc = new Document();

			// Load a document from the specified input file
			doc.LoadFromFile(input);

			// Iterate through the sections in the document
			foreach (Section section in doc.Sections)
			{
				// Iterate through the child objects in the section's body
				foreach (DocumentObject docObj in section.Body.ChildObjects)
				{
					// Check if the current object is a StructureDocumentTag
					if (docObj is StructureDocumentTag)
					{
						// Get the StructureDocumentTag object and its SDTProperties
						StructureDocumentTag stdTag = (StructureDocumentTag)docObj;
						SDTProperties sDTProperties = stdTag.SDTProperties;

					// Set the appearance of the StructureDocumentTag based on its SDTType
					switch (sDTProperties.SDTType)
					{
						case SdtType.Text:
							sDTProperties.Appearance = SdtAppearance.BoundingBox;
							break;
						case SdtType.RichText:
							sDTProperties.Appearance = SdtAppearance.Hidden;
							break;
						case SdtType.Picture:
							sDTProperties.Appearance = SdtAppearance.Tags;
							break;
						case SdtType.CheckBox:
							sDTProperties.Appearance = SdtAppearance.Default;
							break;
					}
				}
			}
			}

			// Specify the output file path
			string output = "SetContentControlAppearance.docx";

			// Save the modified document to the output file
			doc.SaveToFile(output, FileFormat.Docx2013);

			// Dispose the document object
			doc.Dispose();
			
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
