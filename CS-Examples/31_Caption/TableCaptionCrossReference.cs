using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace TableCaptionCrossReference
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
            Document document = new Document();

            //Get the first section
            Section section = document.AddSection();

            //Create a table
            Table table = section.AddTable(true);
            table.ResetCells(2, 3);
            //Add caption to the table
            IParagraph captionParagraph = table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem);

            //Create a bookmark
            string bookmarkName = "Table_1";
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendBookmarkStart(bookmarkName);
            paragraph.AppendBookmarkEnd(bookmarkName);

            //Replace bookmark content
            BookmarksNavigator navigator = new BookmarksNavigator(document);
            navigator.MoveToBookmark(bookmarkName);
            TextBodyPart part = navigator.GetBookmarkContent();
            part.BodyItems.Clear();
            part.BodyItems.Add(captionParagraph);
            navigator.ReplaceBookmarkContent(part);

            //Create cross-reference field to point to bookmark "Table_1"
            Field field = new Field(document);
            field.Type = FieldType.FieldRef;
            field.Code = @"REF Table_1 \p \h";

            //Insert line breaks
            for (int i = 0; i < 3; i++)
            {
                paragraph.AppendBreak(BreakType.LineBreak);
            }

            //Insert field to paragraph
            paragraph = section.AddParagraph();
            TextRange range = paragraph.AppendText("This is a table caption cross-reference, ");
            range.CharacterFormat.FontSize = 14;
            paragraph.ChildObjects.Add(field);

            //Insert FieldSeparator object
            FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
            paragraph.ChildObjects.Add(fieldSeparator);

            //Set display text of the field
            TextRange tr = new TextRange(document);
            tr.Text = "Table 1";
            tr.CharacterFormat.FontSize = 14;
            tr.CharacterFormat.TextColor = System.Drawing.Color.DeepSkyBlue;
            paragraph.ChildObjects.Add(tr);

            //Insert FieldEnd object to mark the end of the field
            FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
            paragraph.ChildObjects.Add(fieldEnd);

            //Update fields
            document.IsUpdateFields = true;

            //Save the file
            string output = "TableCaptionCrossReference.docx";
            document.SaveToFile(output,FileFormat.Docx);

            //Launching the file
            WordDocViewer(output);

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
