using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;

namespace PreserveWordBookmarks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Sample.doc");

            ToPdfParameterList toPdf = new ToPdfParameterList();
            toPdf.CreateWordBookmarks = true;
            toPdf.WordBookmarksTitle = "Bookmark";
            toPdf.WordBookmarksColor = Color.Gray;

            //the event of BookmarkLayout occurs when drawing a bookmark
            document.BookmarkLayout += new Spire.Doc.Documents.Rendering.BookmarkLevelHandler(document_BookmarkLayout);

			//Save the document to a PDF file.
            document.SaveToFile("PreserveBookmarks.pdf", toPdf);
            
			//Launch the file.
            FileViewer("PreserveBookmarks.pdf");
        }
        static void document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args)
        {

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

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
