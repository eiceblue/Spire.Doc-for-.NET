using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace HeaderAndFooter
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

			 //Load the file from disk
			 document.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.docx");

			 //Get the first section
			 Section section = document.Sections[0];

			 //Insert header and footer
			 InsertHeaderAndFooter(section);

			 //Save the file.
			 document.SaveToFile("Sample.docx", FileFormat.Docx);

			 //Dispose the document
			 document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");
        }

        private void InsertHeaderAndFooter(Section section)
        {
            //Get the header
			HeaderFooter header = section.HeadersFooters.Header;

			//Get the footer
			HeaderFooter footer = section.HeadersFooters.Footer;

			// Create a new paragraph for the header and add an image
			Paragraph headerParagraph = header.AddParagraph();
			DocPicture headerPicture = headerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Header.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            DocPicture headerPicture= headerParagraph.AppendPicture(TestUtil.DataPath + "Demo/Header.png");
            */
            // Add text to the header paragraph and set its formatting properties
            TextRange text = headerParagraph.AppendText("Demo of Spire.Doc");
			text.CharacterFormat.FontName = "Arial";
			text.CharacterFormat.FontSize = 10;
			text.CharacterFormat.Italic = true;
            headerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			// Set border properties for the bottom border of the header paragraph
            headerParagraph.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;
			headerParagraph.Format.Borders.Bottom.Space = 0.05F;

			// Set the text wrapping style and alignment properties for the header picture
			headerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
			headerPicture.HorizontalOrigin = HorizontalOrigin.Page;
			headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			headerPicture.VerticalOrigin = VerticalOrigin.Page;
			headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top;

			// Create a new paragraph for the footer and add an image
			Paragraph footerParagraph = footer.AddParagraph();
			DocPicture footerPicture = footerParagraph.AppendPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Footer.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            DocPicture footerPicture = footerParagraph.AppendPicture(TestUtil.DataPath + "Demo/Footer.png");
            */
            // Set the text wrapping style and alignment properties for the footer picture
            footerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
			footerPicture.HorizontalOrigin = HorizontalOrigin.Page;
			footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
			footerPicture.VerticalOrigin = VerticalOrigin.Page;
			footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom;

			// Add fields for page number and total number of pages to the footer paragraph
			footerParagraph.AppendField("page number", FieldType.FieldPage);
			footerParagraph.AppendText(" of ");
			footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
            footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			// Set border properties for the top border of the footer paragraph
            footerParagraph.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
			footerParagraph.Format.Borders.Top.Space = 0.05F;
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
