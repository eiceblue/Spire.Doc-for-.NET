using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace DifferentFirstPage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the document
            string input = @"..\..\..\..\..\..\Data\MultiplePages.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the section and set the property true
            Section section = doc.Sections[0];
            section.PageSetup.DifferentFirstPageHeaderFooter = true;

            //Set the first page header. Here we append a picture in the header
            Paragraph paragraph1 = section.HeadersFooters.FirstPageHeader.AddParagraph();
            paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
            DocPicture headerimage = paragraph1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));

            //Set the first page footer
            Paragraph paragraph2 = section.HeadersFooters.FirstPageFooter.AddParagraph();
            paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            TextRange FF = paragraph2.AppendText("First Page Footer");
            FF.CharacterFormat.FontSize = 10;

            //Set the other header & footer. If you only need the first page header & footer, don't set this
            Paragraph paragraph3 = section.HeadersFooters.Header.AddParagraph();
            paragraph3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            TextRange NH = paragraph3.AppendText("Spire.Doc for .NET");
            NH.CharacterFormat.FontSize = 10;

            Paragraph paragraph4 = section.HeadersFooters.Footer.AddParagraph();
            paragraph4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            TextRange NF = paragraph4.AppendText("E-iceblue");
            NF.CharacterFormat.FontSize = 10;

            //Save and launch document
            string output = "DifferentFirstPage.docx";
            doc.SaveToFile(output, FileFormat.Docx);
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
