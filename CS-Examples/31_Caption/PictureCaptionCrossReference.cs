using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Interface;

namespace PictureCaptionCrossReference
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

            //Create a new section
            Section section = document.AddSection();

            //Add the first paragraph
            Paragraph firstPara = section.AddParagraph();

            //Add the first picture
            Paragraph par1 = section.AddParagraph();
            par1.Format.AfterSpacing = 10;
            DocPicture pic1 = par1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));
            pic1.Height = 120;
            pic1.Width = 120;
            //Add caption to the picture
            CaptionNumberingFormat format = CaptionNumberingFormat.Number;
            IParagraph captionParagraph = pic1.AddCaption("Figure", format, CaptionPosition.BelowItem);
            section.AddParagraph();

            //Add the second picture
            Paragraph par2 = section.AddParagraph();
            DocPicture pic2 = par2.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));
            pic2.Height = 120;
            pic2.Width = 120;
            //Add caption to the picture
            captionParagraph = pic2.AddCaption("Figure", format, CaptionPosition.BelowItem);
            section.AddParagraph();

            //Create a bookmark
            string bookmarkName = "Figure_2";
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

            //Create cross-reference field to point to bookmark "Figure_2"
            Field field = new Field(document);
            field.Type = FieldType.FieldRef;
            field.Code = @"REF Figure_2 \p \h";
            firstPara.ChildObjects.Add(field);
            FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
            firstPara.ChildObjects.Add(fieldSeparator);

            //Set the display text of the field
            TextRange tr = new TextRange(document);
            tr.Text = "Figure 2";
            firstPara.ChildObjects.Add(tr);

            FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
            firstPara.ChildObjects.Add(fieldEnd);

            //Update fields
            document.IsUpdateFields = true;

            //Save the file
            string output = "PictureCaptionCrossReference.docx";
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
