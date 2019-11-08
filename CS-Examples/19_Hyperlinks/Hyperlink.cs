using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Hyperlink
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Open a blank word document as template
            Document document = new Document();
            Section section = document.AddSection();

            //Insert hyperlink
            InsertHyperlink(section);

            //Save doc file.
            document.SaveToFile("Sample.docx", FileFormat.Docx);

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }

        private void InsertHyperlink(Section section)
        {
            Paragraph paragraph
                = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
            paragraph.AppendText("Spire.Doc for .NET \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);

            paragraph = section.AddParagraph();
            paragraph.AppendText("Home page");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            paragraph = section.AddParagraph();
            paragraph.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);

            paragraph = section.AddParagraph();
            paragraph.AppendText("Contact US");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            paragraph = section.AddParagraph();
            paragraph.AppendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.EMailLink);

            paragraph = section.AddParagraph();
            paragraph.AppendText("Forum");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            paragraph = section.AddParagraph();
            paragraph.AppendHyperlink("www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", HyperlinkType.WebLink);

            paragraph = section.AddParagraph();
            paragraph.AppendText("Download Link");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            paragraph = section.AddParagraph();
            paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", "www.e-iceblue.com/Download/download-word-for-net-now.html", HyperlinkType.WebLink);

            paragraph = section.AddParagraph();
            paragraph.AppendText("Insert Link On Image");
            paragraph.ApplyStyle(BuiltinStyle.Heading2);
            paragraph = section.AddParagraph();
            DocPicture picture = paragraph.AppendPicture(System.Drawing.Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));
            paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", picture, HyperlinkType.WebLink);
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
