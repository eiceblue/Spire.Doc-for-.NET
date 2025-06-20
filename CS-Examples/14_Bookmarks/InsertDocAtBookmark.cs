using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertDocAtBookmark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create the first document
            Document document1 = new Document();

            //Load the first document from disk.
            document1.LoadFromFile(@"..\..\..\..\..\..\Data\Bookmark.docx");

            //Create the second document
            Document document2 = new Document();

            //Load the second document from disk.
            document2.LoadFromFile(@"..\..\..\..\..\..\Data\Insert.docx");

            //Get the first section of the first document 
            Section section1 = document1.Sections[0];

            //Locate the bookmark
            BookmarksNavigator bn = new BookmarksNavigator(document1);

            //Find bookmark by name
            bn.MoveToBookmark("Test", true, true);

            //Get bookmarkStart
            BookmarkStart start = bn.CurrentBookmark.BookmarkStart;

            //Get the owner paragraph
            Paragraph para = start.OwnerParagraph;

            //Get the para index
            int index = section1.Body.ChildObjects.IndexOf(para);

            //Loop through the sections
            foreach (Section section2 in document2.Sections)
            {
                foreach (Paragraph paragraph in section2.Paragraphs)
                {
					//Insert the paragraphs of document2
                    section1.Body.ChildObjects.Insert(index++ + 1, paragraph.Clone() as Paragraph);
                }
            }

            //Save the document.
            document1.SaveToFile("Output.docx", FileFormat.Docx);
			
			//Dispose the document
			document1.Dispose();
			document2.Dispose();

            //Launch the Word file.
            WordDocViewer("Output.docx");
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
