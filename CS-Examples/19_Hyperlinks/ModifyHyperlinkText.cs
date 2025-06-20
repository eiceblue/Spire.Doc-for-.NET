using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace ModifyHyperlinkText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
     
			// Specify the input file path for the document containing hyperlinks
			string input = @"..\..\..\..\..\..\Data\Hyperlinks.docx";

			// Create a new Document object
			Document doc = new Document();

			// Load the document from the specified file path
			doc.LoadFromFile(input);

			// Create a list to store the hyperlinks
			List<Field> hyperlinks = new List<Field>();

			// Iterate through the sections in the document
			foreach (Section section in doc.Sections)
			{
				// Iterate through the child objects in the body of the section
				foreach (DocumentObject sec in section.Body.ChildObjects)
				{
					// Check if the child object is a paragraph
					if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
					{
						// Iterate through the child objects in the paragraph
						foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
						{
							// Check if the child object is a field
							if (para.DocumentObjectType == DocumentObjectType.Field)
							{
								// Cast the child object to a Field
								Field field = para as Field;

								// Check if the field is a hyperlink
								if (field.Type == FieldType.FieldHyperlink)
								{
									// Add the field to the list of hyperlinks
									hyperlinks.Add(field);
								}
							}
						}
					}
				}
			}

			// Modify the text of the first hyperlink field
			hyperlinks[0].FieldText = "Spire.Doc component";

			// Specify the output file path for the modified document
			string output = "ModifyText.docx";

			// Save the modified document to the output file path in DOCX format
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document object to free up resources
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
