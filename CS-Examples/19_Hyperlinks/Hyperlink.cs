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
			// Create a new Document object
			Document document = new Document();

			// Add a section to the document
			Section section = document.AddSection();

			// Call the InsertHyperlink method to insert hyperlinks in the section
			InsertHyperlink(section);

			// Save the document to a file named "Sample.docx" in DOCX format
			document.SaveToFile("Sample.docx", FileFormat.Docx);

			// Dispose the document object to free up resources
			document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("Sample.docx");


        }

		// Define the InsertHyperlink method
		private void InsertHyperlink(Section section)
		{
			// Add a paragraph to the section, or get the first paragraph if it exists
			Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
			
			// Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Spire.Doc for .NET \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			// Add a new paragraph to the section
			paragraph = section.AddParagraph();
			
			// Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Home page");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);
			
			// Add a hyperlink to the paragraph with the specified URL and display text
			paragraph = section.AddParagraph();
			paragraph.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);

			// Add a new paragraph to the section
			paragraph = section.AddParagraph();
			
			// Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Contact US");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);
			
			// Add a hyperlink to the paragraph with the specified email address and display text
			paragraph = section.AddParagraph();
			paragraph.AppendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.EMailLink);

			// Add a new paragraph to the section
			paragraph = section.AddParagraph();
			
			// Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Forum");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);
			
			// Add a hyperlink to the paragraph with the specified URL and display text
			paragraph = section.AddParagraph();
			paragraph.AppendHyperlink("www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", HyperlinkType.WebLink);

			// Add a new paragraph to the section
			paragraph = section.AddParagraph();
			
			// Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Download Link");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);
			
			// Add a hyperlink to the paragraph with the specified URL and display text
			paragraph = section.AddParagraph();
			paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", "www.e-iceblue.com/Download/download-word-for-net-now.html", HyperlinkType.WebLink);

			// Add a new paragraph to the section
			paragraph = section.AddParagraph();
			
			// Set the text content and apply a built-in style to the paragraph
			paragraph.AppendText("Insert Link On Image");
			paragraph.ApplyStyle(BuiltinStyle.Heading2);
			
			// Add an image to the paragraph and append a hyperlink to it with the specified URL and link type
			paragraph = section.AddParagraph();
			DocPicture picture = paragraph.AppendPicture(System.Drawing.Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture picture = paragraph.AppendPicture(@"..\..\..\..\..\..\Data\Spire.Doc.png");
            */
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
