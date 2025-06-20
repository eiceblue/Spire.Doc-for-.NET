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

namespace SetLocaleForField
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
			Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_2.docx");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Add a paragraph to the section
			Paragraph par = section.AddParagraph();

			// Append a date field to the paragraph
			Field field = par.AppendField("DocDate", FieldType.FieldDate);

			// Set the locale ID to Russian (1049) for the first character range in the field
			(field.OwnerParagraph.ChildObjects[0] as TextRange).CharacterFormat.LocaleIdASCII = 1049;

			// Set the field text to "2019-10-10"
			field.FieldText = "2019-10-10";

			// Enable automatic update of fields in the document
			document.IsUpdateFields = true;

			// Specify the output file name
			string result = "result.docx";

			// Save the modified document to a new file
			document.SaveToFile(result, FileFormat.Docx);

			// Dispose of the document object
			document.Dispose();
			
            //Launch result file
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
