using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace CreateBookmark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
             // Create a Document object
			 Document document = new Document();

			 // Add a Section to the document
			 Section section = document.AddSection();

			 // Call the CreateBookmark method to create bookmarks in the section
			 CreateBookmark(section);

			 // Save the document
			 document.SaveToFile("Output.docx", FileFormat.Docx);

			 // Dispose the document
			 document.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
        }

        private void CreateBookmark(Section section)
        {
            // Add a Paragraph to the section
			Paragraph paragraph = section.AddParagraph();

			// Add text with formatting and make it italic
			TextRange txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark in a Word document.");
			txtRange.CharacterFormat.Italic = true;

			// Add an empty paragraph for spacing
			section.AddParagraph();

			// Add another paragraph with heading style and colored text
			paragraph = section.AddParagraph();
			txtRange = paragraph.AppendText("Simple Create Bookmark.");
			txtRange.CharacterFormat.TextColor = Color.CornflowerBlue;
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			// Add an empty paragraph for spacing
			section.AddParagraph();

			// Add a paragraph with a bookmark
			paragraph = section.AddParagraph();
			paragraph.AppendBookmarkStart("SimpleCreateBookmark");
			paragraph.AppendText("This is a simple bookmark.");
			paragraph.AppendBookmarkEnd("SimpleCreateBookmark");

			// Add an empty paragraph for spacing
			section.AddParagraph();

			// Add another paragraph with heading style and colored text
			paragraph = section.AddParagraph();
			txtRange = paragraph.AppendText("Nested Create Bookmark.");
			txtRange.CharacterFormat.TextColor = Color.CornflowerBlue;
			paragraph.ApplyStyle(BuiltinStyle.Heading2);

			// Add an empty paragraph for spacing
			section.AddParagraph();

			// Add a paragraph with nested bookmarks
			paragraph = section.AddParagraph();
			paragraph.AppendBookmarkStart("Root");
			txtRange = paragraph.AppendText(" This is Root data ");
			txtRange.CharacterFormat.Italic = true;
			paragraph.AppendBookmarkStart("NestedLevel1");
			txtRange = paragraph.AppendText(" This is Nested Level1 ");
			txtRange.CharacterFormat.Italic = true;
			txtRange.CharacterFormat.TextColor = Color.DarkSlateGray;
			paragraph.AppendBookmarkStart("NestedLevel2");
			txtRange = paragraph.AppendText(" This is Nested Level2 ");
			txtRange.CharacterFormat.Italic = true;
			txtRange.CharacterFormat.TextColor = Color.DimGray;
			paragraph.AppendBookmarkEnd("NestedLevel2");
			paragraph.AppendBookmarkEnd("NestedLevel1");
			paragraph.AppendBookmarkEnd("Root");

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
