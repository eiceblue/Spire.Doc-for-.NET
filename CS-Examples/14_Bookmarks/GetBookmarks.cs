using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Text;
using System.IO;

namespace GetBookmarks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
            Document document = new Document();
			
			//Load the document from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Bookmarks.docx");

            //Get the bookmark by index.
            Bookmark bookmark1 = document.Bookmarks[0];

            //Get the bookmark by name.
            Bookmark bookmark2 = document.Bookmarks["Test2"];

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set string format for displaying
            string result = string.Format("The bookmark obtained by index is " + bookmark1.Name + ".\r\nThe bookmark obtained by name is " + bookmark2.Name + ".\n");

            //Add result string to StringBuilder
            content.AppendLine(result);

            //Save them to a txt file
            File.WriteAllText("Bookmarks.txt", content.ToString());
			
			// Dispose the document
			document.Dispose();

            //Launch the file
            FileViewer("Bookmarks.txt");
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
