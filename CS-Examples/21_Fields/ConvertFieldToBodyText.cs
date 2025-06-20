using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ConvertFieldToBodyText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
			// Create a new Document object to store the source document
			Document sourceDocument = new Document();

			// Load the source document from a file
			sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\TextInputField.docx");

			// Iterate through each form field in the first section of the document's body
			foreach (FormField field in sourceDocument.Sections[0].Body.FormFields)
			{
				// Check if the form field is of type FieldFormTextInput
				if (field.Type == FieldType.FieldFormTextInput)
				{
					// Get the owner paragraph of the form field
					Paragraph paragraph = field.OwnerParagraph;

					// Initialize variables for start and end index of bookmark objects
					int startIndex = 0;
					int endIndex = 0;

					// Create a TextRange object using the source document
					TextRange textRange = new TextRange(sourceDocument);

					// Set the text of the TextRange to the text of the paragraph
					textRange.Text = paragraph.Text;

					// Iterate through each child object of the paragraph
					foreach (DocumentObject obj in paragraph.ChildObjects)
					{
						// Check if the child object is a BookmarkStart object
						if (obj.DocumentObjectType == DocumentObjectType.BookmarkStart)
						{
							// Store the index of the BookmarkStart object
							startIndex = paragraph.ChildObjects.IndexOf(obj);
						}

						// Check if the child object is a BookmarkEnd object
						if (obj.DocumentObjectType == DocumentObjectType.BookmarkEnd)
						{
							// Store the index of the BookmarkEnd object
							endIndex = paragraph.ChildObjects.IndexOf(obj);
						}
					}

					// Remove the form fields or child objects between the start and end index
					for (int i = endIndex; i > startIndex; i--)
					{
						if (paragraph.ChildObjects[i] is TextFormField)
						{
							// Remove the TextFormField object
							TextFormField textFormField = paragraph.ChildObjects[i] as TextFormField;
							paragraph.ChildObjects.Remove(textFormField);
						}
						else
						{
							// Remove other child objects
							paragraph.ChildObjects.RemoveAt(i);
						}
					}

					// Insert the modified TextRange at the start index of the paragraph
					paragraph.ChildObjects.Insert(startIndex, textRange);

					// Exit the loop after processing the first FieldFormTextInput
					break;
				}
			}

			// Save the modified document to a new file
			sourceDocument.SaveToFile("Output.docx", FileFormat.Docx);

			// Dispose the source document object
			sourceDocument.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
