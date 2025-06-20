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

namespace InsertPageRefField
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
			Document document = new Document(@"..\..\..\..\..\..\Data\PageRef.docx");

			// Get the last section of the document
			Section section = document.LastSection;

			// Add a paragraph to the section
			Paragraph par = section.AddParagraph();

			// Append a page reference field with the specified name and type
			Field field = par.AppendField("pageRef", FieldType.FieldPageRef);

			// Set the code for the field with the specified parameters
			field.Code = "PAGEREF bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT";

			// Enable the automatic update of fields in the document
			document.IsUpdateFields = true;

			// Save the modified document to a file with the specified name
			document.SaveToFile("result.docx", FileFormat.Docx);

			// Dispose of the document object to free up resources
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
