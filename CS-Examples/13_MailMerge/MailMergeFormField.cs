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
namespace MailMergeFormField
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Path to the input document
			string input = @"..\..\..\..\..\..\Data\MailMergeFormField.doc";
			
			// Create a new Document instance
			Document document = new Document();

			// Load the document from the specified file
			document.LoadFromFile(input);

			// Define the field names for the mail merge
			string[] fieldNames = new string[] { "Contact Name", "Fax", "Date", "Urgent", "Share", "Submit", "Body" };

			// Define the field values for the mail merge
			string[] fieldValues = new string[] { "John Smith", "+1 (69) 123456", DateTime.Now.Date.ToString(),
				"Yes","No","Yes",
				"<b>It's very urgent. Please deal with it ASAP. </b>" };

			// Subscribe to the MergeField event
			document.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);

			// Execute the mail merge using the field names and values
			document.MailMerge.Execute(fieldNames, fieldValues);

			// Specify the output file name for the merged document
			string result = "MailMergeFormField_out.docx";

			// Save the merged document to the specified file in Docx format
			document.SaveToFile(result, FileFormat.Docx);

			// Dispose the document object
			document.Dispose();
				
            WordViewer(result);
        }

        void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            if (args.FieldValue.ToString() == "Yes")
			{
				// Get the checkbox name from the field name
				string checkBoxName = args.FieldName;

				// Get the owner paragraph of the current merge field
				Paragraph para = args.CurrentMergeField.OwnerParagraph;

				// Get the index of the current merge field within its parent paragraph
				int index = para.ChildObjects.IndexOf(args.CurrentMergeField);

				// Create a new CheckBoxFormField
				CheckBoxFormField field = para.AppendField(checkBoxName, FieldType.FieldFormCheckBox) as CheckBoxFormField;

				// Insert the new checkbox field at the same index as the current merge field
				para.ChildObjects.Insert(index, field);

				// Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField);

				// Set the checkbox field as checked
				field.Checked = true;
			}
			
			if (args.FieldValue.ToString() == "No")
			{
				// Get the checkbox name from the field name
				string checkBoxName = args.FieldName;

				// Get the owner paragraph of the current merge field
				Paragraph para = args.CurrentMergeField.OwnerParagraph;

				// Get the index of the current merge field within its parent paragraph
				int index = para.ChildObjects.IndexOf(args.CurrentMergeField);

				// Create a new CheckBoxFormField
				CheckBoxFormField field = para.AppendField(checkBoxName, FieldType.FieldFormCheckBox) as CheckBoxFormField;

				// Insert the new checkbox field at the same index as the current merge field
				para.ChildObjects.Insert(index, field);

				// Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField);

				// Set the checkbox field as unchecked
				field.Checked = false;
			}
		   
			if (args.FieldName == "Body")
			{
				// Get the owner paragraph of the current merge field
				Paragraph para = args.CurrentMergeField.OwnerParagraph;

				// Append the HTML content as plain text to the paragraph
				para.AppendHTML(args.FieldValue.ToString());

				// Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField);
			}

			if (args.FieldName == "Date")
			{
				// Get the text input name from the field name
				string textInputName = args.FieldName;

				// Get the owner paragraph of the current merge field
				Paragraph para = args.CurrentMergeField.OwnerParagraph;

				// Create a new TextFormField
				TextFormField field = para.AppendField(textInputName, FieldType.FieldFormTextInput) as TextFormField;

				// Remove the current merge field from the paragraph
				para.ChildObjects.Remove(args.CurrentMergeField);

				// Set the text value for the text input field
				field.Text = args.FieldValue.ToString();
			}
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
