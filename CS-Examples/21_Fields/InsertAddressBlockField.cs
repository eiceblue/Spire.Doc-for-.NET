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

namespace InsertAddressBlockField
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

			// Add a new paragraph to the section
			Paragraph par = section.AddParagraph();

			// Append a field with type "AddressBlock" to the paragraph
			Field field = par.AppendField("ADDRESSBLOCK", FieldType.FieldAddressBlock);

			// Set the code for the field, including additional options and formatting
			field.Code = "ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"";

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
