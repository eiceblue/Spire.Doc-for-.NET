using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Linq;
using Spire.Doc.Reporting;
using System.Collections;
using System.Collections.Generic;
using Spire.Doc.Interface;
namespace ExecuteConditionalField
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
			Document doc = new Document();

			// Add a Section to the document
			Section section = doc.AddSection();

			// Add a Paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Create and add the first IF field
			CreateIFField1(doc, paragraph);

			// Add another paragraph to the section
			paragraph = section.AddParagraph();

			// Create and add the second IF field
			CreateIFField2(doc, paragraph);

			// Set up field names and values for mail merge
			string[] fieldName = { "Count", "Age" };
			string[] fieldValue = { "2", "30" };

			// Execute the mail merge
			doc.MailMerge.Execute(fieldName, fieldValue);

			// Set IsUpdateFields property to true
			doc.IsUpdateFields = true;

			// Save the document 
			string result = "ExecuteConditionalField_out.docx";
			doc.SaveToFile(result, FileFormat.Docx);

			// Dispose the document object
			doc.Dispose();
            WordViewer(result);
        }
        private void CreateIFField1(Document document, Paragraph paragraph)
        {
            // Create a new IfField object
			IfField ifField = new IfField(document);

			// Set the type and code of the IfField
			ifField.Type = FieldType.FieldIf;
			ifField.Code = "IF ";

			// Add the IfField to the paragraph
			paragraph.Items.Add(ifField);

			// Append the fields and text to the paragraph
			paragraph.AppendField("Count", FieldType.FieldMergeField);
			paragraph.AppendText(" > ");
			paragraph.AppendText("\"1\" ");
			paragraph.AppendText("\"Greater than one\" ");
			paragraph.AppendText("\"Less than one\"");

			// Create and add the end field mark
			IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
			(end as FieldMark).Type = FieldMarkType.FieldEnd;
			paragraph.Items.Add(end);

			// Set the end field mark for the IfField
			ifField.End = end as FieldMark;
        }

        private void CreateIFField2(Document document, Paragraph paragraph)
        {
            // Create a new IfField object
			IfField ifField = new IfField(document);

			// Set the type and code of the IfField
			ifField.Type = FieldType.FieldIf;
			ifField.Code = "IF ";

			// Add the IfField to the paragraph
			paragraph.Items.Add(ifField);

			// Append the fields and text to the paragraph
			paragraph.AppendField("Age", FieldType.FieldMergeField);
			paragraph.AppendText(" > ");
			paragraph.AppendText("\"50\" ");
			paragraph.AppendText("\"The old man\" ");
			paragraph.AppendText("\"The young man\"");

			// Create and add the end field mark
			IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
			(end as FieldMark).Type = FieldMarkType.FieldEnd;
			paragraph.Items.Add(end);

			// Set the end field mark for the IfField
			ifField.End = end as FieldMark;
        }
        private void WordViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
