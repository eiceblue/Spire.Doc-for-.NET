using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ImageHeaderAndFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\Template.docx";

			//Create a word document
			Document doc = new Document();

			//Load the document from disk
			doc.LoadFromFile(input);

			//Get the header of the first page
			HeaderFooter header = doc.Sections[0].HeadersFooters.Header;

			//Add a paragraph for the header
			Paragraph paragraph = header.AddParagraph();

			//Set the format of the paragraph
            paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			//Append a picture in the paragraph
			DocPicture headerimage = paragraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));
			headerimage.VerticalAlignment = ShapeVerticalAlignment.Bottom;

			//Get the footer of the first section
			HeaderFooter footer = doc.Sections[0].HeadersFooters.Footer;

			//Add a paragraph for the footer
			Paragraph paragraph2 = footer.AddParagraph();

			//Set the format of the paragraph
            paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

			//Append a picture in the paragraph
			paragraph2.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\logo.png"));

			//Append text in the paragraph and set its character format
			TextRange TR = paragraph2.AppendText("Copyright © 2013 e-iceblue. All Rights Reserved.");
			TR.CharacterFormat.FontName = "Arial";
			TR.CharacterFormat.FontSize = 10;
			TR.CharacterFormat.TextColor = Color.Black;

			//Save the document
			string output = "ImageHeaderAndFooter.docx";
			doc.SaveToFile(output, FileFormat.Docx);

			//Dispose the document
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
