using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Fields.OMath;

namespace ConvertEqToOfficeMath
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
			Document document = new Document();

			// Load the document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\EQ.docx");

			// Get the first paragraph of the first section in the document
			Paragraph paragraph = document.Sections[0].Paragraphs[0];

			// Iterate through the child objects of the paragraph
			for (int i = 0; i < paragraph.ChildObjects.Count; i++)
			{
				// Get the current document object
				DocumentObject documentObject = paragraph.ChildObjects[i];

				// Check if the document object is a field of type Equation
				if (documentObject is Field && ((Field)documentObject).Type == FieldType.FieldEquation)
				{
					// Convert the field to an OfficeMath object
					OfficeMath officeMath = OfficeMath.FromEqField((Field)documentObject);

					// If conversion is successful, replace the field with the OfficeMath object
					if (officeMath != null)
					{
						paragraph.ChildObjects.Remove(documentObject);
						paragraph.ChildObjects.Insert(i, officeMath);
					}
				}
			}

			// Save the modified document to a new file
			document.SaveToFile("ConvertEqToOfficeMath.docx", FileFormat.Docx);

			// Dispose of the document object
			document.Dispose();

            //Launch the Word file.
            WordDocViewer("ConvertEqToOfficeMath.docx");
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
