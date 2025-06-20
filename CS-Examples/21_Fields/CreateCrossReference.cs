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
         
			// Create a new Document object
			Document document = new Document();

			// Add a section to the document
			Section section = document.AddSection();

			// Add a paragraph to the section and append a bookmark with the specified name
			Paragraph paragraph = section.AddParagraph();
			paragraph.AppendBookmarkStart("MyBookmark");
			paragraph.AppendText("Text inside a bookmark");
			paragraph.AppendBookmarkEnd("MyBookmark");

			// Add line breaks to the paragraph
			for (int i = 0; i < 4; i++)
			{
				paragraph.AppendBreak(BreakType.LineBreak);
			}

			// Create a new Field object for referencing the bookmark
			Field field = new Field(document);
			field.Type = FieldType.FieldRef;
			field.Code = @"REF MyBookmark \p \h";

			// Add a new paragraph to the section and append text and the field
			paragraph = section.AddParagraph();
			paragraph.AppendText("For more information, see ");
			paragraph.ChildObjects.Add(field);

			// Add a field separator to the paragraph
			FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
			paragraph.ChildObjects.Add(fieldSeparator);

			// Create a TextRange object and set its text
			TextRange tr = new TextRange(document);
			tr.Text = "above";
			paragraph.ChildObjects.Add(tr);

			// Add a field end mark to the paragraph
			FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
			paragraph.ChildObjects.Add(fieldEnd);

			// Specify the file name for the result document
			String result = "Result-CreateCrossReferenceToBookmark.docx";

			// Save the document to a file
			document.SaveToFile(result, FileFormat.Docx2013);

			// Dispose the document object
			document.Dispose();

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
