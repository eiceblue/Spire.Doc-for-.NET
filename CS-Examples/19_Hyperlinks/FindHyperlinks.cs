using System;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;

namespace FindHyperlinks
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

			// Create a list to store the hyperlinks and a variable to hold the text of the hyperlinks
			List<Field> hyperlinks = new List<Field>();
			string hyperlinksText = null;

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
									
									// Append the field's text to the hyperlinksText variable
									hyperlinksText += field.FieldText + "\r\n";
								}
							}
						}
					}
				}
			}

			// Specify the output file path for the generated text file
			string output = "HyperlinksText.txt";

			// Write the hyperlinks text to the output file
			File.WriteAllText(output, hyperlinksText);

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
