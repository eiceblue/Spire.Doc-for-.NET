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
            string input = @"..\..\..\..\..\..\Data\MultiplePages.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Get the section
			Section section = doc.Sections[0];

			//specify that the current section has a different header/footer for the first page
			section.PageSetup.DifferentFirstPageHeaderFooter = true;

			//Set the first page header. Here we append a picture in the header
			Paragraph paragraph1 = section.HeadersFooters.FirstPageHeader.AddParagraph();

			//Set horizontal alignment for the paragraph
            paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			//Append a picture
			DocPicture headerimage = paragraph1.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));

			//Set the first page footer
			Paragraph paragraph2 = section.HeadersFooters.FirstPageFooter.AddParagraph();

			//Set horizontal alignment for the paragraph
            paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Append text
			TextRange FF = paragraph2.AppendText("First Page Footer");

			//Set font size
			FF.CharacterFormat.FontSize = 10;

			//Set the other header & footer. If you only need the first page header & footer, don't set this
			Paragraph paragraph3 = section.HeadersFooters.Header.AddParagraph();

			//Set horizontal alignment for the paragraph
            paragraph3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Append text
			TextRange NH = paragraph3.AppendText("Spire.Doc for .NET");

			//Set font size
			NH.CharacterFormat.FontSize = 10;

			//Add a paragraph
			Paragraph paragraph4 = section.HeadersFooters.Footer.AddParagraph();

			//Set horizontal alignment for the paragraph
            paragraph4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			//Append text
			TextRange NF = paragraph4.AppendText("E-iceblue");

			//Set font size
			NF.CharacterFormat.FontSize = 10;

			//Save the document
			string output = "DifferentFirstPage.docx";
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
