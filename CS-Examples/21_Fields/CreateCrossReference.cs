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

namespace CreateCrossReference
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

            //Create a bookmark.
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendBookmarkStart("MyBookmark");
            paragraph.AppendText("Text inside a bookmark");
            paragraph.AppendBookmarkEnd("MyBookmark");

            //Insert line breaks.
            for (int i = 0; i < 4; i++)
            {
                paragraph.AppendBreak(BreakType.LineBreak);
            }

            //Create a cross-reference field, and link it to bookmark.                    
            Field field = new Field(document);
            field.Type = FieldType.FieldRef;
            field.Code = @"REF MyBookmark \p \h";

            //Insert field to paragraph.
            paragraph = section.AddParagraph();
            paragraph.AppendText("For more information, see ");
            paragraph.ChildObjects.Add(field);

            //Insert FieldSeparator object.
            FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
            paragraph.ChildObjects.Add(fieldSeparator);

            //Set display text of the field.
            TextRange tr = new TextRange(document);
            tr.Text = "above";
            paragraph.ChildObjects.Add(tr);

            //Insert FieldEnd object to mark the end of the field.
            FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
            paragraph.ChildObjects.Add(fieldEnd);

            String result = "Result-CreateCrossReferenceToBookmark.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
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
