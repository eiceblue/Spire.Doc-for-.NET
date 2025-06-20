using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace CreateIFField
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

			// Create an IF field and add it to the paragraph
			CreateIfField(document, paragraph);

			// Set field name and value for mail merge
			string[] fieldName = { "Count" };
			string[] fieldValue = { "2" };

			// Execute the mail merge
			document.MailMerge.Execute(fieldName, fieldValue);

			// Enable field update after mail merge
			document.IsUpdateFields = true;

			// Specify the file name for saving the document
			String result = "Result-CreateAnIFField.docx";

			// Save the document to a file
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose the document object
			document.Dispose();

            //Launch the file.
            WordDocViewer(result);
        }

  
		// Method to create an IF field
		static void CreateIfField(Document document, Paragraph paragraph)
		{
			// Create a new IF field
			IfField ifField = new IfField(document);
			ifField.Type = FieldType.FieldIf;
			ifField.Code = "IF ";

			// Add the IF field to the paragraph
			paragraph.Items.Add(ifField);

			// Add the merge field and condition to the paragraph
			paragraph.AppendField("Count", FieldType.FieldMergeField);
			paragraph.AppendText(" > ");
			paragraph.AppendText("\"100\" ");
			paragraph.AppendText("\"Thanks\" ");
			paragraph.AppendText("\"The minimum order is 100 units\"");

			// Create the end mark of the IF field and add it to the paragraph
			IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
			(end as FieldMark).Type = FieldMarkType.FieldEnd;
			paragraph.Items.Add(end);

			// Set the end mark of the IF field
			ifField.End = end as FieldMark;
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
