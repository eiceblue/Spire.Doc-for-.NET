using System;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ExtractBookmarkText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Document
            string input = @"..\..\..\..\..\..\Data\BookmarkTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Creates a BookmarkNavigator instance to access the bookmark
            BookmarksNavigator navigator = new BookmarksNavigator(doc);
            //Locate a specific bookmark by bookmark name
            navigator.MoveToBookmark("Content");
            TextBodyPart textBodyPart = navigator.GetBookmarkContent();

            //Iterate through the items in the bookmark content to get the text
            string text = null;
            foreach (var item in textBodyPart.BodyItems)
            {
                if (item is Paragraph)
                {
                    foreach (var childObject in (item as Paragraph).ChildObjects)
                    {
                        if (childObject is TextRange)
                        {
                            text += (childObject as TextRange).Text;
                        }
                    }
                }
            }

            //Save to TXT File and launch it
            string output = "ExtractBookmarkText.txt";
            File.WriteAllText(output, text);
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
