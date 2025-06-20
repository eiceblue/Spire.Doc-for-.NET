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
            string input = @"..\..\..\..\..\..\Data\BookmarkTemplate.docx";

			//Create a word document
			Document doc = new Document();

			//Load the file from disk
			doc.LoadFromFile(input);

			//Creates a BookmarkNavigator instance to access the bookmark
			BookmarksNavigator navigator = new BookmarksNavigator(doc);

			//Locate a specific bookmark by bookmark name
			navigator.MoveToBookmark("Content");

			//Get the bookmark content
			TextBodyPart textBodyPart = navigator.GetBookmarkContent();

			//Define a variable to store the text
			string text = null;

			//Iterate through the items in the bookmark content to get the text
			foreach (var item in textBodyPart.BodyItems)
			{
				if (item is Paragraph)
				{
					//Iterate through the child objects of the paragraph
					foreach (var childObject in (item as Paragraph).ChildObjects)
					{
						if (childObject is TextRange)
						{
							//Append the text
							text += (childObject as TextRange).Text;
						}
					}
				}
			}


			string output = "ExtractBookmarkText.txt";

			//Save to TXT File and launch it
			File.WriteAllText(output, text);

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
