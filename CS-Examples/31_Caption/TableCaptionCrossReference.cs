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
            // Create a new instance of Document
            Document document = new Document();

            // Add a section to the document
            Section section = document.AddSection();

            // Add a table to the section with 2 rows and 3 columns
            Table table = section.AddTable(true);
            table.ResetCells(2, 3);

            // Add a caption to the table with the "Table" label, numbering format as "Number", and position below the table
            IParagraph captionParagraph = table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem);

            // Add a bookmark at the specified location
            string bookmarkName = "Table_1";
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendBookmarkStart(bookmarkName);
            paragraph.AppendBookmarkEnd(bookmarkName);

            // Navigate to the bookmark and replace its content with the caption paragraph
            BookmarksNavigator navigator = new BookmarksNavigator(document);
            navigator.MoveToBookmark(bookmarkName);
            TextBodyPart part = navigator.GetBookmarkContent();
            part.BodyItems.Clear();
            part.BodyItems.Add(captionParagraph);
            navigator.ReplaceBookmarkContent(part);

            // Create a cross-reference field for the bookmark
            Field field = new Field(document);
            field.Type = FieldType.FieldRef;
            field.Code = @"REF Table_1 \p \h";

            // Add line breaks before the next paragraph
            for (int i = 0; i < 3; i++)
            {
                paragraph.AppendBreak(BreakType.LineBreak);
            }

            // Add a new paragraph for the caption cross-reference
            paragraph = section.AddParagraph();

            // Add text to the paragraph
            TextRange range = paragraph.AppendText("This is a table caption cross-reference, ");
            range.CharacterFormat.FontSize = 14;

            // Add the field for referencing the table caption
            paragraph.ChildObjects.Add(field);

            // Add a field separator
            FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
            paragraph.ChildObjects.Add(fieldSeparator);

            // Add the text "Table 1" as the reference text
            TextRange tr = new TextRange(document);
            tr.Text = "Table 1";
            tr.CharacterFormat.FontSize = 14;
            tr.CharacterFormat.TextColor = System.Drawing.Color.DeepSkyBlue;
            paragraph.ChildObjects.Add(tr);

            // Add a field end mark
            FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
            paragraph.ChildObjects.Add(fieldEnd);

            // Enable field updating in the document
            document.IsUpdateFields = true;

            // Specify the output file name and format (Docx)
            string output = "TableCaptionCrossReference.docx";
            document.SaveToFile(output, FileFormat.Docx);

            // Dispose of the document object when finished using it
            document.Dispose();

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
