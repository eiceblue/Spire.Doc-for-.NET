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
            Document doc = new Document();
            //Add a new section 
            Section section = doc.AddSection();
            //Add a new paragraph for a section 
            Paragraph paragraph = section.AddParagraph();

            CreateIFField1(doc, paragraph);
            paragraph = section.AddParagraph();
            CreateIFField2(doc, paragraph);

            string[] fieldName = { "Count","Age" };
            string[] fieldValue = { "2","30" };

            doc.MailMerge.Execute(fieldName, fieldValue);
            doc.IsUpdateFields = true;

            doc.SaveToFile("sample.docx", FileFormat.Docx);

            string result = "ExecuteConditionalField_out.docx";
            doc.SaveToFile(result, FileFormat.Docx);
            WordViewer(result);
        }
        private void CreateIFField1(Document document, Paragraph paragraph)
        {
            IfField ifField = new IfField(document);
            ifField.Type = FieldType.FieldIf;
            ifField.Code = "IF ";
            paragraph.Items.Add(ifField);

            paragraph.AppendField("Count", FieldType.FieldMergeField);
            paragraph.AppendText(" > ");
            paragraph.AppendText("\"1\" ");
            paragraph.AppendText("\"Greater than one\" ");
            paragraph.AppendText("\"Less than one\"");

            IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
            (end as FieldMark).Type = FieldMarkType.FieldEnd;
            paragraph.Items.Add(end);

            ifField.End = end as FieldMark;
        }

        private void CreateIFField2(Document document, Paragraph paragraph)
        {
            IfField ifField = new IfField(document);
            ifField.Type = FieldType.FieldIf;
            ifField.Code = "IF ";
            paragraph.Items.Add(ifField);

            paragraph.AppendField("Age", FieldType.FieldMergeField);
            paragraph.AppendText(" > ");
            paragraph.AppendText("\"50\" ");
            paragraph.AppendText("\"The old man\" ");
            paragraph.AppendText("\"The young man\"");

            IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
            (end as FieldMark).Type = FieldMarkType.FieldEnd;
            paragraph.Items.Add(end);

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
