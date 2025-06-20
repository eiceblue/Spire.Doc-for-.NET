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
            // Create a new instance of the Document class
			Document document = new Document();

			// Load a Word document from the specified file path
			document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Sample.doc");

			// Create a ToPdfParameterList object to specify conversion parameters
			ToPdfParameterList toPdf = new ToPdfParameterList();

			// Set the 'CreateWordBookmarks' parameter to true to preserve bookmarks in the converted PDF
			toPdf.CreateWordBookmarks = true;

			// Set the 'WordBookmarksTitle' parameter to specify the title of the bookmarks in the PDF
			toPdf.WordBookmarksTitle = "Bookmark";

			// Set the 'WordBookmarksColor' parameter to specify the color of the bookmarks in the PDF
			toPdf.WordBookmarksColor = Color.Gray;

			// Attach an event handler to the BookmarkLayout event of the document
			document.BookmarkLayout += new Spire.Doc.Documents.Rendering.BookmarkLevelHandler(document_BookmarkLayout);

			// Save the document as a PDF with the specified conversion parameters
			document.SaveToFile("PreserveBookmarks.pdf", toPdf);
            
			//Launch the file.
            FileViewer("PreserveBookmarks.pdf");
        }
		
        // Define the event handler for the BookmarkLayout event
		static void document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args)
		{
			// Customize the appearance of bookmarks based on their level
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
