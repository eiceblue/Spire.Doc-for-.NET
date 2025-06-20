using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetCultureForField
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

			// Add a section to the document
			Section section = document.AddSection();

			// Add a paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Add text to the paragraph
			paragraph.AppendText("Add Date Field: ");

			// Append a date field to the paragraph and set its format
			Field field1 = paragraph.AppendField("Date1", FieldType.FieldDate) as Field;
			field1.Code = @"DATE  \@" + "\"yyyy\\MM\\dd\"";

			// Add a new paragraph to the section
			Paragraph newParagraph = section.AddParagraph();

			// Add text to the new paragraph
			newParagraph.AppendText("Add Date Field with setting French Culture: ");

			// Append a date field to the new paragraph and set its format
			Field field2 = newParagraph.AppendField("\"\\@\"dd MMMM yyyy", FieldType.FieldDate);
			field2.CharacterFormat.LocaleIdASCII = 1036;

			// Enable automatic update of fields in the document
			document.IsUpdateFields = true;

			// Save the document to a file
			document.SaveToFile("Output.docx", FileFormat.Docx);

			// Dispose of the document object
			document.Dispose();

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
