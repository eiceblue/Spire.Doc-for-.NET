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

namespace InsertAdvanceField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
			// Load the document from a specified file path
			Document document = new Document(@"..\..\..\..\..\..\Data\SampleB_2.docx");

			// Get the first section of the document
			Section section = document.Sections[0];

			// Add a paragraph to the section
			Paragraph par = section.AddParagraph();

			// Append a field with the specified type and text
			Field field = par.AppendField("Field", FieldType.FieldAdvance);

			// Set the code for the field using the specified parameters
			field.Code = "ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 ";

			// Enable the automatic update of fields in the document
			document.IsUpdateFields = true;

			// Save the modified document to a file with the specified name
			String result = "result.docx";
			document.SaveToFile(result, FileFormat.Docx);

			// Dispose of the document object to free up resources
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
