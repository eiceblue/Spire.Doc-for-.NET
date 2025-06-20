using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace CreateNestedField
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
			Paragraph paragraph = section.AddParagraph();

			// Create the outer IF field and add it to the paragraph
			IfField ifField = new IfField(document);
			ifField.Type = FieldType.FieldIf;
			ifField.Code = "IF ";
			paragraph.Items.Add(ifField);

			// Create the inner IF field and add it to the paragraph
			IfField ifField2 = new IfField(document);
			ifField2.Type = FieldType.FieldIf;
			ifField2.Code = "IF ";
			paragraph.ChildObjects.Add(ifField2);
			paragraph.Items.Add(ifField2);
			paragraph.AppendText("\"200\" < \"50\"   \"200\" \"50\" ");

			// Create the end mark for the inner IF field and add it to the paragraph
			IParagraphBase embeddedEnd = document.CreateParagraphItem(ParagraphItemType.FieldMark);
			(embeddedEnd as FieldMark).Type = FieldMarkType.FieldEnd;
			paragraph.Items.Add(embeddedEnd);
			ifField2.End = embeddedEnd as FieldMark;

			// Append additional text and create the end mark for the outer IF field
			paragraph.AppendText(" > ");
			paragraph.AppendText("\"100\" ");
			paragraph.AppendText("\"Thanks\" ");
			paragraph.AppendText("\"The minimum order is 100 units\"");
			IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
			(end as FieldMark).Type = FieldMarkType.FieldEnd;
			paragraph.Items.Add(end);
			ifField.End = end as FieldMark;

			// Enable field update
			document.IsUpdateFields = true;

			// Specify the file name for saving the document
			String result = "CreateNestedField_output.docx";

			// Save the document to a file
			document.SaveToFile(result, FileFormat.Docx);

			// Dispose the document object
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
