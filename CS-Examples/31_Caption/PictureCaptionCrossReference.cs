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
            // Create a new instance of Document
            Document document = new Document();

            // Add a new section to the document
            Section section = document.AddSection();

            // Add a paragraph to the section for the cross-reference
            Paragraph firstPara = section.AddParagraph();

            // Add another paragraph to the section
            Paragraph par1 = section.AddParagraph();
            par1.Format.AfterSpacing = 10;

            // Append an image (picture) to the paragraph from the specified file path
            DocPicture pic1 = par1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            DocPicture pic1 = par1.AppendPicture(@"..\..\..\..\..\..\Data\Spire.Doc.png");
            */
            pic1.Height = 120;
            pic1.Width = 120;

            // Set the caption numbering format to "Number" and add a caption below the picture
            CaptionNumberingFormat format = CaptionNumberingFormat.Number;
            IParagraph captionParagraph = pic1.AddCaption("Figure", format, CaptionPosition.BelowItem);

            // Add another paragraph to the section
            Paragraph par2 = section.AddParagraph();

            // Append another image (picture) to the paragraph from the specified file path
            DocPicture pic2 = par2.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture pic2 = par2.AppendPicture(@"..\..\..\..\..\..\Data\Word.png");
            */
            pic2.Height = 120;
            pic2.Width = 120;

            // Add a caption below the second picture
            captionParagraph = pic2.AddCaption("Figure", format, CaptionPosition.BelowItem);

            // Add a bookmark at the specified location
            string bookmarkName = "Figure_2";
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
            field.Code = @"REF Figure_2 \p \h";
            firstPara.ChildObjects.Add(field);
            FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
            firstPara.ChildObjects.Add(fieldSeparator);

            // Add the text "Figure 2" as the reference text
            TextRange tr = new TextRange(document);
            tr.Text = "Figure 2";
            firstPara.ChildObjects.Add(tr);

            FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
            firstPara.ChildObjects.Add(fieldEnd);

            // Enable field updating in the document
            document.IsUpdateFields = true;

            // Specify the output file name and format (Docx)
            string output = "PictureCaptionCrossReference.docx";
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
