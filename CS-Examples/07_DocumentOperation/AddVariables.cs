using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AddVariables
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

			// Append a field with the text "A1" and field type FieldDocVariable to the paragraph
			paragraph.AppendField("A1", FieldType.FieldDocVariable);

			// Add a variable named "A1" with a value of "12" to the document's variables collection
			document.Variables.Add("A1", "12");

			// Set the IsUpdateFields property of the document to true, enabling field updates
			document.IsUpdateFields = true;

			// Specify the file name for the saved document
			string result = "Result-AddVariables.docx";

			// Save the document to a file in DOCX format (using Word 2013 format)
			document.SaveToFile(result, FileFormat.Docx2013);

			// Release the resources used by the document
			document.Dispose();

            //Launch the MS Word file.
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
