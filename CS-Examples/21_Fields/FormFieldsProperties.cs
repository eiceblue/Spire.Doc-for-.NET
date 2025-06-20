using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Collections;
using System.Text;

namespace FormFieldsProperties
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
			// Load the document from a file
			Document document = new Document(@"..\..\..\..\..\..\Data\FillFormField.doc");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Get the second form field in the section
			FormField formField = section.Body.FormFields[1];

			// Check if the form field is a text input field
			if (formField.Type == FieldType.FieldFormTextInput)
			{
				// Set the text of the form field
				formField.Text = "My name is " + formField.Name;

				// Customize the text formatting of the form field
				formField.CharacterFormat.TextColor = Color.Red;
				formField.CharacterFormat.Italic = true;
			}

			// Save the modified document to a file
			document.SaveToFile("result.docx", FileFormat.Docx);

			// Dispose the document object
			document.Dispose();
			
            //Launch result file
            WordDocViewer("result.docx");

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
