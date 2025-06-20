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

namespace InsertMergeField
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

			// Append a merge field with the specified name and type
			MergeField field = par.AppendField("MyFieldName", FieldType.FieldMergeField) as MergeField;

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
