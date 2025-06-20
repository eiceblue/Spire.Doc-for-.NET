using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace StartFromFormField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class to store the source document.
            Document sourceDocument = new Document();

            // Load the source document from the specified file path.
            sourceDocument.LoadFromFile(@"..\..\..\..\..\..\Data\TextInputField.docx");

            // Create a new instance of the Document class to store the destination document.
            Document destinationDoc = new Document();

            // Add a new section to the destination document.
            Section section = destinationDoc.AddSection();

            // Initialize the index variable to 0.
            int index = 0;

            // Iterate through each form field in the body of the source document's first section.
            foreach (FormField field in sourceDocument.Sections[0].Body.FormFields)
            {
                // Check if the form field is of type FieldFormTextInput.
                if (field.Type == FieldType.FieldFormTextInput)
                {
                    // Get the paragraph that contains the form field.
                    Paragraph paragraph = field.OwnerParagraph;

                    // Find the index of the paragraph within the child objects of the source document's body.
                    index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(paragraph);

                    // Exit the loop after finding the first form text input field.
                    break;
                }
            }

            // Copy three consecutive child objects starting from the found index from the source document's body to the destination document's section.
            for (int i = index; i < index + 3; i++)
            {
                // Clone the child object at the current index.
                DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

                // Add the cloned child object to the body of the destination document's section.
                section.Body.ChildObjects.Add(doobj);
            }

            // Save the destination document to a new file named "FromFormField.docx" in DOCX format.
            destinationDoc.SaveToFile("FromFormField.docx", FileFormat.Docx);

            // Dispose of the source document object to release resources.
            sourceDocument.Dispose();

            // Dispose of the destination document object to release resources.
            destinationDoc.Dispose();

            //Launch the Word file.
            WordDocViewer("FromFormField.docx");
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
