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
            //Create Word document.
            Document document = new Document();

            //Add a new section.
            Section section = document.AddSection();

            //Add a new paragraph.
            Paragraph paragraph = section.AddParagraph();

            // Define a method of creating an IF Field.
            CreateIfField(document, paragraph);

            //Define merged data.
            string[] fieldName = { "Count" };
            string[] fieldValue = { "2" };

            //Merge data into the IF Field.
            document.MailMerge.Execute(fieldName, fieldValue);

            //Update all fields in the document.
            document.IsUpdateFields = true;

            String result = "Result-CreateAnIFField.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the file.
            WordDocViewer(result);
        }

        //Create the IF Field like:{IF { MERGEFIELD Count } > "100" "Thanks" " The minimum order is 100 units "}
        static void CreateIfField(Document document, Paragraph paragraph)
        {
            IfField ifField = new IfField(document);
            ifField.Type = FieldType.FieldIf;
            ifField.Code = "IF ";

            paragraph.Items.Add(ifField);
            paragraph.AppendField("Count", FieldType.FieldMergeField);
            paragraph.AppendText(" > ");
            paragraph.AppendText("\"100\" ");
            paragraph.AppendText("\"Thanks\" ");
            paragraph.AppendText("\"The minimum order is 100 units\"");

            IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
            (end as FieldMark).Type = FieldMarkType.FieldEnd;
            paragraph.Items.Add(end);
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
