using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AddTCField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
     
			// Create a new Document object
			Document document = new Document();

			// Add a new section to the document
			Section section = document.AddSection();

			// Add a new paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Append a TC (Table of Contents) field to the paragraph with the specified entry text
			Field field = paragraph.AppendField("TC", FieldType.FieldTOCEntry);
			field.Code = @"TC " + "\"Entry Text\"" + " \\f" + " t";

			// Save the document to a file with the specified file name and format (Docx)
			document.SaveToFile("AddTCField.docx", FileFormat.Docx);

			// Dispose the Document object to free resources
			document.Dispose();

            //Launch result file and please set "Show all formatting marks" to display the field 
            WordDocViewer("AddTCField.docx");
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
