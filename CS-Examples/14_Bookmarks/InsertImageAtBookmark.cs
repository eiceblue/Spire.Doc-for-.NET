using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace InsertImageAtBookmark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
			string input = @"..\..\..\..\..\..\Data\Bookmark.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Create an instance of BookmarksNavigator
			BookmarksNavigator bn = new BookmarksNavigator(doc);

			//Find a bookmark named Test
			bn.MoveToBookmark("Test", true, true);

			//Add a section
			Section section0 = doc.AddSection();

			//Add a paragraph for the section
			Paragraph paragraph = section0.AddParagraph();
			
			//Load an image
			Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Word.png");

			//Add a picture into the paragraph
			DocPicture picture = paragraph.AppendPicture(image);
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
             DocPicture picture = paragraph.AppendPicture(@"..\..\..\..\..\..\Data\Word.png");
            */

            //Add the paragraph at the position of bookmark
            bn.InsertParagraph(paragraph);

			//Remove the section0
			doc.Sections.Remove(section0);

			//Save the document
			string output = "InsertImageAtBookmark.docx";
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
