using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;

namespace ConvertIfFieldToText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
			// Create a new Document object and load the document from a file
			Document document = new Document(@"..\..\..\..\..\..\Data\IfFieldSample.docx");

			// Get the collection of fields in the document
			FieldCollection fields = document.Fields;

			// Iterate through each field in the collection
			for (int i = 0; i < fields.Count; i++)
			{
				// Get the current field
				Field field = fields[i];

				// Check if the field is of type FieldIf
				if (field.Type == FieldType.FieldIf)
				{
					// Cast the field as TextRange to access its properties
					TextRange original = field as TextRange;

					// Get the text of the field
					string text = field.FieldText;

					// Create a new TextRange object with the document and set its text to the field text
					TextRange textRange = new TextRange(document);
					textRange.Text = text;

					// Set the font name and size of the new text range to match the original field
					textRange.CharacterFormat.FontName = original.CharacterFormat.FontName;
					textRange.CharacterFormat.FontSize = original.CharacterFormat.FontSize;

					// Get the owner paragraph of the field
					Paragraph par = field.OwnerParagraph;

					// Get the index of the field within its owner paragraph
					int index = par.ChildObjects.IndexOf(field);

					// Remove the field from its owner paragraph
					par.ChildObjects.RemoveAt(index);

					// Insert the new text range at the index of the field within its owner paragraph
					par.ChildObjects.Insert(index, textRange);
				}
			}

			// Specify the file name for the result document
			String result = "result.docx";

			// Save the modified document to a new file
			document.SaveToFile(result, FileFormat.Docx);

			// Dispose the document object
			document.Dispose();

            //Launch the Word file
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
