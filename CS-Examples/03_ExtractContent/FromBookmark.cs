using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;

namespace FromBookmark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create the source document
            Document sourcedocument = new Document();

            //Load the source document from disk.
            sourcedocument.LoadFromFile(@"..\..\..\..\..\..\Data\Bookmark.docx");

            //Create a destination document
            Document destinationDoc = new Document();

            //Add a section for destination document
            Section section = destinationDoc.AddSection();

            //Add a paragraph for destination document
            Paragraph paragraph = section.AddParagraph();

            //Locate the bookmark in source document
            BookmarksNavigator navigator = new BookmarksNavigator(sourcedocument);

            //Find bookmark by name
            navigator.MoveToBookmark("Test", true, true);

            //get text body part
            TextBodyPart textBodyPart = navigator.GetBookmarkContent();

            //Create a TextRange type list
            List<TextRange> list = new List<TextRange>();

            //Traverse the items of text body
            foreach (var item in textBodyPart.BodyItems)
            {
                //if it is paragraph
                if (item is Paragraph)
                {
                    //Traverse the ChildObjects of the paragraph
                    foreach (var childObject in (item as Paragraph).ChildObjects)
                    {
                        //if it is TextRange
                        if (childObject is TextRange)
                        {
                            //Add it into list
                            TextRange range = childObject as TextRange;
                            list.Add(range);
                        }
                    }
                }
            }

            //Add the extract content to destinationDoc document
            for (int m = 0; m < list.Count; m++)
            {
                paragraph.Items.Add(list[m].Clone());
            }

            //Save the document.
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

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
