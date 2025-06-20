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
            // Create a new Document object to represent the source document.
            Document sourcedocument = new Document();

            // Load the Word document from the specified file path.
            sourcedocument.LoadFromFile(@"..\..\..\..\..\..\Data\Bookmark.docx");

            // Create a new Document object to represent the destination document.
            Document destinationDoc = new Document();

            // Add a section to the destination document.
            Section section = destinationDoc.AddSection();

            // Add a paragraph to the section.
            Paragraph paragraph = section.AddParagraph();

            // Create a BookmarksNavigator object using the source document.
            BookmarksNavigator navigator = new BookmarksNavigator(sourcedocument);

            // Move the navigator to the bookmark with the specified name.
            navigator.MoveToBookmark("Test", true, true);

            // Get the content of the bookmark as a TextBodyPart.
            TextBodyPart textBodyPart = navigator.GetBookmarkContent();

            // Create a list to store the TextRanges extracted from the bookmark.
            List<TextRange> list = new List<TextRange>();

            // Iterate over each body item in the TextBodyPart.
            foreach (var item in textBodyPart.BodyItems)
            {
                // Check if the body item is a Paragraph.
                if (item is Paragraph)
                {
                    // Iterate over each child object in the Paragraph.
                    foreach (var childObject in (item as Paragraph).ChildObjects)
                    {
                        // Check if the child object is a TextRange.
                        if (childObject is TextRange)
                        {
                            // Cast the child object to TextRange and add it to the list.
                            TextRange range = childObject as TextRange;
                            list.Add(range);
                        }
                    }
                }
            }

            // Copy the TextRanges from the list to the destination document's paragraph.
            for (int m = 0; m < list.Count; m++)
            {
                paragraph.Items.Add(list[m].Clone());
            }

            // Save the destination document to a file.
            destinationDoc.SaveToFile("Output.docx", FileFormat.Docx);

            // Dispose of the source and destination documents to free up resources.
            sourcedocument.Dispose();
            destinationDoc.Dispose();

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
