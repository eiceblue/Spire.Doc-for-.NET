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
            // Create a new document
			Document doc = new Document();

			// Add a section to the document
			Section section = doc.AddSection();

			// Set the page border properties
            section.PageSetup.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Wave;
			section.PageSetup.Borders.Color = Color.Green;
			section.PageSetup.Borders.Left.Space = 20;
			section.PageSetup.Borders.Right.Space = 20;

			// Add a header paragraph to the section
			Paragraph paragraph1 = section.HeadersFooters.Header.AddParagraph();

			//Set horizontal alignment for the paragraph
            paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			//Append text
			TextRange headerText = paragraph1.AppendText("Header isn't included in page border");

			//Set the character format for the text
			headerText.CharacterFormat.FontName = "Calibri";
			headerText.CharacterFormat.FontSize = 20;
			headerText.CharacterFormat.Bold = true;

			// Add a footer paragraph to the section
			Paragraph paragraph2 = section.HeadersFooters.Footer.AddParagraph();

			//Set horizontal alignment for the paragraph
            paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

			//Append text
			TextRange footerText = paragraph2.AppendText("Footer is included in page border");

			//Set the character format for the text
			footerText.CharacterFormat.FontName = "Calibri";
			footerText.CharacterFormat.FontSize = 20;
			footerText.CharacterFormat.Bold = true;

			// Configure page setup properties
			section.PageSetup.PageBorderIncludeHeader = false;
			section.PageSetup.HeaderDistance = 40;
			section.PageSetup.PageBorderIncludeFooter = true;
			section.PageSetup.FooterDistance = 40;

			// Save the document to a file
			string output = "PageBorderSurround.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			// Dispose the document
			doc.Dispose();
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
