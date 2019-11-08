using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CreateBookmarkForTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document.
            Document document = new Document();

            //Add a new section.
            Section section = document.AddSection();

            //Create bookmark for a table
            CreateBookmarkForTable(document, section);

            String result = "Output.docx";
            //Save the document.
            document.SaveToFile(result, FileFormat.Docx);

            //Launch the Word file.
            WordDocViewer(result);
        }

        private void CreateBookmarkForTable(Document doc, Section section)
        {
            //Add a paragraph
            Paragraph paragraph = section.AddParagraph();

            //Append text for added paragraph
            TextRange txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark for a table in a Word document.");

            //Set the font in italic
            txtRange.CharacterFormat.Italic = true;

            //Append bookmark start
            paragraph.AppendBookmarkStart("CreateBookmark");

            //Append bookmark end
            paragraph.AppendBookmarkEnd("CreateBookmark");

            //Add table
            Table table = section.AddTable(true);

            //Set the number of rows and columns
            table.ResetCells(2, 2);

            //Append text for table cells
            TextRange range = table[0, 0].AddParagraph().AppendText("sampleA");
            range = table[0, 1].AddParagraph().AppendText("sampleB");
            range = table[1, 0].AddParagraph().AppendText("120");
            range = table[1, 1].AddParagraph().AppendText("260");

            //Get the bookmark by index.
            Bookmark bookmark = doc.Bookmarks[0];

            //Get the name of bookmark.
            String bookmarkName = bookmark.Name;

            //Locate the bookmark by name.
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            navigator.MoveToBookmark(bookmarkName);

            //Add table to TextBodyPart
            TextBodyPart part = navigator.GetBookmarkContent();
            part.BodyItems.Add(table);

            //Replace bookmark cotent with table
            navigator.ReplaceBookmarkContent(part);
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
