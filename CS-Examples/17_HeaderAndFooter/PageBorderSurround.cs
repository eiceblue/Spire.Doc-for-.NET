using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace PageBorderSurround
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document
            Document doc = new Document();
            Section section = doc.AddSection();

            //Add a sample page border to the document
            section.PageSetup.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Wave;
            section.PageSetup.Borders.Color = Color.Green;
            section.PageSetup.Borders.Left.Space = 20;
            section.PageSetup.Borders.Right.Space = 20;

            //Add a header and set its format
            Paragraph paragraph1 = section.HeadersFooters.Header.AddParagraph();
            paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
            TextRange headerText = paragraph1.AppendText("Header isn't included in page border");
            headerText.CharacterFormat.FontName = "Calibri";
            headerText.CharacterFormat.FontSize = 20;
            headerText.CharacterFormat.Bold = true;

            //Add a footer and set its format
            Paragraph paragraph2 = section.HeadersFooters.Footer.AddParagraph();
            paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
            TextRange footerText = paragraph2.AppendText("Footer is included in page border");
            footerText.CharacterFormat.FontName = "Calibri";
            footerText.CharacterFormat.FontSize = 20;
            footerText.CharacterFormat.Bold = true;

            //Set the header not included in the page border while the footer included
            section.PageSetup.PageBorderIncludeHeader = false;
            section.PageSetup.HeaderDistance = 40;
            section.PageSetup.PageBorderIncludeFooter = true;
            section.PageSetup.FooterDistance = 40;

            //Save and launch document
            string output = "PageBorderSurround.docx";
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
