using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SetBookmarkColor
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
            string input = @"..\..\..\..\..\..\Data\BookmarkTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Create an instance of ToPdfParameterList
            ToPdfParameterList toPdf = new ToPdfParameterList();

            //Set CreateWordBookmarks to true to use word bookmarks when create the bookmarks
            toPdf.CreateWordBookmarks = true;

            //Set the title of word bookmarks
            toPdf.WordBookmarksTitle = "Changed bookmark";

            //Set the text color of word bookmarks
            toPdf.WordBookmarksColor = Color.Gray;

            //Call the event document_BookmarkLayout when drawing a bookmark
            doc.BookmarkLayout += new Spire.Doc.Documents.Rendering.BookmarkLevelHandler(document_BookmarkLayout);

            //Save and launch document
            string output = "SetBookmarkColor.pdf";
            doc.SaveToFile(output, toPdf);
            Viewer(output);
        }
        static void document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args)
        {
            //set the different color for different levels of bookmarks
            if (args.BookmarkLevel.Level == 2)
            {
                args.BookmarkLevel.Color = Color.Red;
                args.BookmarkLevel.Style = BookmarkTextStyle.Bold;
            }
            else if (args.BookmarkLevel.Level == 3)
            {
                args.BookmarkLevel.Color = Color.Gray;
                args.BookmarkLevel.Style = BookmarkTextStyle.Italic;
            }
            else
            {
                args.BookmarkLevel.Color = Color.Green;
                args.BookmarkLevel.Style = BookmarkTextStyle.Regular;
            }
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
