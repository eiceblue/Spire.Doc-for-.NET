using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Encrypt
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

            Section section = document.AddSection();

            //page setup
            SetPage(section);

            //insert header and footer
            InsertHeaderAndFooter(section);

            //add content
            InsertContent(section);

            //encrypt document with password specified by textBox1
            document.Encrypt(this.textBox1.Text);

            //Save doc file.
            document.SaveToFile("Sample.doc",FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");


        }

        private void InsertHeaderAndFooter(Section section)
        {
            HeaderFooter header = section.HeadersFooters.Header;
            HeaderFooter footer = section.HeadersFooters.Footer;

            //insert picture and text to header
            Paragraph headerParagraph = header.AddParagraph();
            DocPicture headerPicture
                = headerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Header.png"));

            //header text
            TextRange text = headerParagraph.AppendText("Demo of Spire.Doc");
            text.CharacterFormat.FontName = "Arial";
            text.CharacterFormat.FontSize = 10;
            text.CharacterFormat.Italic = true;
            headerParagraph.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Right;

            //border
            headerParagraph.Format.Borders.Bottom.BorderType
                = Spire.Doc.Documents.BorderStyle.Single;
            headerParagraph.Format.Borders.Bottom.Space = 0.05F;


            //header picture layout - text wrapping
            headerPicture.TextWrappingStyle = TextWrappingStyle.Behind;

            //header picture layout - position
            headerPicture.HorizontalOrigin = HorizontalOrigin.Page;
            headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            headerPicture.VerticalOrigin = VerticalOrigin.Page;
            headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top;

            //insert picture to footer
            Paragraph footerParagraph = footer.AddParagraph();
            DocPicture footerPicture
                = footerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Footer.png"));

            //footer picture layout
            footerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
            footerPicture.HorizontalOrigin = HorizontalOrigin.Page;
            footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            footerPicture.VerticalOrigin = VerticalOrigin.Page;
            footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom;

            //insert page number
            footerParagraph.AppendField("page number", FieldType.FieldPage);
            footerParagraph.AppendText(" of ");
            footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
            footerParagraph.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Right;

            //border
            footerParagraph.Format.Borders.Top.BorderType
                = Spire.Doc.Documents.BorderStyle.Single;
            footerParagraph.Format.Borders.Top.Space = 0.05F;
        }

        private void SetPage(Section section)
        {
            //the unit of all measures below is point, 1point = 0.3528 mm
            section.PageSetup.PageSize = PageSize.A4;
            section.PageSetup.Margins.Top = 72f;
            section.PageSetup.Margins.Bottom = 72f;
            section.PageSetup.Margins.Left = 89.85f;
            section.PageSetup.Margins.Right = 89.85f;
        }

        private void InsertContent(Section section)
        {
            //title
            Paragraph paragraph = section.AddParagraph();
            TextRange title = paragraph.AppendText("Summary of Science");
            title.CharacterFormat.Bold = true;
            title.CharacterFormat.FontName = "Arial";
            title.CharacterFormat.FontSize = 14;
            paragraph.Format.HorizontalAlignment
                = Spire.Doc.Documents.HorizontalAlignment.Center;
            paragraph.Format.AfterSpacing = 10;

            //style
            ParagraphStyle style1 = new ParagraphStyle(section.Document);
            style1.Name = "style1";
            style1.CharacterFormat.FontName = "Arial";
            style1.CharacterFormat.FontSize = 9;
            style1.ParagraphFormat.LineSpacing = 1.5F * 12F;
            style1.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            section.Document.Styles.Add(style1);

            ParagraphStyle style2 = new ParagraphStyle(section.Document);
            style2.Name = "style2";
            style2.ApplyBaseStyle(style1.Name);
            style2.CharacterFormat.Font = new Font("Arial", 10f);
            section.Document.Styles.Add(style2);

            paragraph = section.AddParagraph();
            paragraph.AppendText("(All text and pictures are from ");
            String link = "http://en.wikipedia.org/wiki/Science";
            paragraph.AppendHyperlink(link, "Wikipedia", HyperlinkType.WebLink);
            paragraph.AppendText(", the free encyclopedia)");
            paragraph.ApplyStyle(style1.Name);

            Paragraph paragraph1 = section.AddParagraph();
            String str1
                = "Science (from the Latin scientia, meaning \"knowledge\") "
                + "is an enterprise that builds and organizes knowledge in the form "
                + "of testable explanations and predictions about the natural world. "
                + "An older meaning still in use today is that of Aristotle, "
                + "for whom scientific knowledge was a body of reliable knowledge "
                + "that can be logically and rationally explained "
                + "(see \"History and etymology\" section below).";
            paragraph1.AppendText(str1);

            //Insert a picture in the right of the paragraph1
            DocPicture picture
                = paragraph1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Wikipedia_Science.png"));
            picture.TextWrappingStyle = TextWrappingStyle.Square;
            picture.TextWrappingType = TextWrappingType.Left;
            picture.VerticalOrigin = VerticalOrigin.Paragraph;
            picture.VerticalPosition = 0;
            picture.HorizontalOrigin = HorizontalOrigin.Column;
            picture.HorizontalAlignment = ShapeHorizontalAlignment.Right;

            paragraph1.ApplyStyle(style2.Name);

            Paragraph paragraph2 = section.AddParagraph();
            String str2
                = "Since classical antiquity science as a type of knowledge was closely linked "
                + "to philosophy, the way of life dedicated to discovering such knowledge. "
                + "And into early modern times the two words, \"science\" and \"philosophy\", "
                + "were sometimes used interchangeably in the English language. "
                + "By the 17th century, \"natural philosophy\" "
                + "(which is today called \"natural science\") could be considered separately "
                + "from \"philosophy\" in general. But \"science\" continued to also be used "
                + "in a broad sense denoting reliable knowledge about a topic, in the same way "
                + "it is still used in modern terms such as library science or political science.";
            paragraph2.AppendText(str2);
            paragraph2.ApplyStyle(style2.Name);

            Paragraph paragraph3 = section.AddParagraph();
            String str3
                = "The more narrow sense of \"science\" that is common today developed as a part "
                + "of science became a distinct enterprise of defining \"laws of nature\", "
                + "based on early examples such as Kepler's laws, Galileo's laws, and Newton's "
                + "laws of motion. In this period it became more common to refer to natural "
                + "philosophy as  \"natural science\". Over the course of the 19th century, the word "
                + "\"science\" became increasingly associated with the disciplined study of the "
                + "natural world including physics, chemistry, geology and biology. This sometimes "
                + "left the study of human thought and society in a linguistic limbo, which was "
                + "resolved by classifying these areas of academic study as social science. "
                + "Similarly, several other major areas of disciplined study and knowledge "
                + "exist today under the general rubric of \"science\", such as formal science "
                + "and applied science.";
            paragraph3.AppendText(str3);
            paragraph3.ApplyStyle(style2.Name);
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
