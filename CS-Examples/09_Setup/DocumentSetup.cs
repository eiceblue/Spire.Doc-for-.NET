using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace DocumentSetup
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

            document.BuiltinDocumentProperties.Title = "Document Demo Document";
            document.BuiltinDocumentProperties.Subject = "demo";
            document.BuiltinDocumentProperties.Author = "James";
            document.BuiltinDocumentProperties.Company = "e-iceblue";
            document.BuiltinDocumentProperties.Manager = "Jakson";
            document.BuiltinDocumentProperties.Category = "Doc Demos";
            document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";
            document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

            Section section = document.Sections[0];
            section.PageSetup.Margins.Top = 72f;
            section.PageSetup.Margins.Bottom = 72f;
            section.PageSetup.Margins.Left = 89.85f;
            section.PageSetup.Margins.Right = 89.85f;

            String p1
                = "Microsoft Word is a word processor designed by Microsoft. "
                + "It was first released in 1983 under the name Multi-Tool Word for Xenix systems. "
                + "Subsequent versions were later written for several other platforms including "
                + "IBM PCs running DOS (1983), the Apple Macintosh (1984), the AT&T Unix PC (1985), "
                + "Atari ST (1986), SCO UNIX, OS/2, and Microsoft Windows (1989). ";
            String p2
                = "Microsoft Office Word instead of merely Microsoft Word. "
                + "The 2010 version appears to be branded as Microsoft Word, "
                + "once again. The current versions are Microsoft Word 2010 for Windows and 2008 for Mac.";
            Paragraph paragraph
                = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
            paragraph.AppendText(p1).CharacterFormat.FontSize = 14;
            section.AddParagraph().AppendText(p2).CharacterFormat.FontSize = 14;


            //Save doc file.
            document.SaveToFile("Sample.doc",FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");


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
