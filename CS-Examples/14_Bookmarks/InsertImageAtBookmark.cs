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
            //Load the document
            string input = @"..\..\..\..\..\..\Data\Bookmark.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Create an instance of BookmarksNavigator
            BookmarksNavigator bn = new BookmarksNavigator(doc);

            //Find a bookmark named Test
            bn.MoveToBookmark("Test", true, true);

            //Add a section
            Section section0 = doc.AddSection();

            //Add a paragraph for the section
            Paragraph paragraph = section0.AddParagraph();
            Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Word.png");

            //Add a picture into the paragraph
            DocPicture picture = paragraph.AppendPicture(image);

            //Add the paragraph at the position of bookmark
            bn.InsertParagraph(paragraph);

            //Remove the section0
            doc.Sections.Remove(section0);

            //Save and launch document
            string output = "InsertImageAtBookmark.docx";
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
