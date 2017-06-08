using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

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
            Document document = new Document(@"..\..\..\..\..\..\Data\Blank.doc");

            InsertHyberlink(document.Sections[0]);

            //Save doc file.
            document.SaveToFile("Sample.doc",FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");


        }

        private void InsertHyberlink(Section section)
        {
            Paragraph paragraph
                = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
            paragraph.AppendText("Spire.XLS for .NET \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
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
