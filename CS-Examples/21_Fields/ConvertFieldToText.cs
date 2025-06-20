using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Fields;
using System;
using System.Windows.Forms;

namespace ConvertFieldToText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
			// Create a new Document object to store the document
			Document document = new Document();

			// Load the document from a file
			document.LoadFromFile(@"..\..\..\..\..\..\Data\Fields.docx");

			// Get the collection of fields in the document
			FieldCollection fields = document.Fields;
			int count = fields.Count;

			// Iterate through each field in the collection
			for (int i = 0; i < count; i++)
			{
				// Get the first field in the collection
				Field field = fields[0];

				// Get the text of the field
				string s = field.FieldText;

				// Get the index of the field within its owner paragraph
				int index = field.OwnerParagraph.ChildObjects.IndexOf(field);

				// Create a TextRange object with the document and set its text to the field text
				TextRange textRange = new TextRange(document);
				textRange.Text = s;
				
				// Set the font size of the text range
				textRange.CharacterFormat.FontSize = 24f;

				// Insert the text range at the index of the field within its owner paragraph
				field.OwnerParagraph.ChildObjects.Insert(index, textRange);

				// Remove the field from its owner paragraph
				field.OwnerParagraph.ChildObjects.Remove(field);
			}

			// Save the modified document to a new file
			document.SaveToFile("ConvertFieldToText.docx", FileFormat.Docx);

			// Dispose the document object
			document.Dispose();

            //Launching the Word file.
            WordDocViewer("ConvertFieldToText.docx");


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
