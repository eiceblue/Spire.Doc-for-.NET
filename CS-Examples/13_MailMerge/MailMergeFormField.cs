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

            String input = @"..\..\..\..\..\..\Data\MailMergeFormField.doc";
            //Create word document
            Document document = new Document();
            document.LoadFromFile(input);

            string[] fieldNames = new string[] { "Contact Name", "Fax", "Date", "Urgent", "Share", "Submit","Body" };

            string[] fieldValues = new string[] { "John Smith", "+1 (69) 123456", DateTime.Now.Date.ToString(),
                "Yes","No","Yes",             
                "<b>It's very urgent. Please deal with it ASAP. </b>" };

            document.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);
            document.MailMerge.Execute(fieldNames, fieldValues);
            string result = "MailMergeFormField_out.docx";
            document.SaveToFile(result, FileFormat.Docx);
            WordViewer(result);
        }

        void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            if (args.FieldValue == "Yes")
            {
                //Create a checkbox name
               string checkBoxName = args.FieldName;
               Paragraph para =  args.CurrentMergeField.OwnerParagraph;
               int index = para.ChildObjects.IndexOf(args.CurrentMergeField);
                // Insert a check box.
               CheckBoxFormField field = para.AppendField(checkBoxName, FieldType.FieldFormCheckBox) as CheckBoxFormField;
               para.ChildObjects.Insert(index, field);
               para.ChildObjects.Remove(args.CurrentMergeField);
               field.Checked = true;
               
            }
            if (args.FieldValue == "No")
            {
                //Create a checkbox name
                string checkBoxName = args.FieldName;
                Paragraph para = args.CurrentMergeField.OwnerParagraph;
                int index = para.ChildObjects.IndexOf(args.CurrentMergeField);
                // Insert a check box.
                CheckBoxFormField field = para.AppendField(checkBoxName, FieldType.FieldFormCheckBox) as CheckBoxFormField;
                para.ChildObjects.Insert(index, field);
                para.ChildObjects.Remove(args.CurrentMergeField);
                field.Checked = false;
            }
            // Insert html during mail merge.
            if (args.FieldName == "Body")
            {
                Paragraph para = args.CurrentMergeField.OwnerParagraph;
                para.AppendHTML(args.FieldValue.ToString());
                para.ChildObjects.Remove(args.CurrentMergeField);
            }

            // Insert text input form field.
            if (args.FieldName == "Date")
            {
                string textInputName = args.FieldName;
                Paragraph para = args.CurrentMergeField.OwnerParagraph;
                TextFormField field = para.AppendField(textInputName, FieldType.FieldFormTextInput) as TextFormField;
                para.ChildObjects.Remove(args.CurrentMergeField);
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
