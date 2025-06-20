using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ToPDFAndCreateBookmarks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Define the input file path
			string inputFile = "../../../../../../../Data/BookmarkTemplate.docx";

			// Define the output file name
			string outFile = "ToPDFAndCreateBookmarks_out.pdf";

			// Create a new Document instance
			Document document = new Document();

			// Load the document from the specified input file
			document.LoadFromFile(inputFile);

			// Create a ToPdfParameterList instance to configure PDF conversion options
			ToPdfParameterList parames = new ToPdfParameterList();

			// Enable the creation of bookmarks in the resulting PDF
			parames.CreateWordBookmarks = true;

			// Choose whether to create bookmarks using headings (true) or not (false)
			parames.CreateWordBookmarksUsingHeadings = false;
			// Uncomment this line to enable creating bookmarks using headings
			//parames.CreateWordBookmarksUsingHeadings = true; 

			// Save the document as a PDF file with the specified output file name and conversion parameters
			document.SaveToFile(outFile, parames);

			// Dispose the Document object after use
			document.Dispose();
            WordDocViewer(outFile);
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
