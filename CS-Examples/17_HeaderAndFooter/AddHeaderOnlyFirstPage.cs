using System;
using System.Windows.Forms;
using Spire.Doc;

namespace AddHeaderOnlyFirstPage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string input = @"..\..\..\..\..\..\Data\HeaderAndFooter.docx";

			//Create a word document
			Document doc1 = new Document();

			//Load the source file 
			doc1.LoadFromFile(input);

			//Get the header from the first section
			HeaderFooter header = doc1.Sections[0].HeadersFooters.Header;

			input = @"..\..\..\..\..\..\Data\MultiplePages.docx";

			//Create another word document
			Document doc2 = new Document();

			//Load the destination file
			doc2.LoadFromFile(input);

			//Get the first page header of the destination document
			HeaderFooter firstPageHeader = doc2.Sections[0].HeadersFooters.FirstPageHeader;

			//Loop the sections of doc2
			foreach (Section section in doc2.Sections)
			{
				//Specify that the current section has a different header/footer for the first page
				section.PageSetup.DifferentFirstPageHeaderFooter = true;
			}

			//Removes all child objects in firstPageHeader
			firstPageHeader.Paragraphs.Clear();

			//Loop through the child objects of the header
			foreach (DocumentObject obj in header.ChildObjects)
			{
				//Add all child objects of the header to firstPageHeader
				firstPageHeader.ChildObjects.Add(obj.Clone());
			}

			//Save and launch the file
			string resultfile = "AddHeaderOnlyFirstPage.docx";
			doc2.SaveToFile(resultfile, FileFormat.Docx);

			// Dispose the document
			doc1.Dispose();
			doc2.Dispose();
            Viewer(resultfile);
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
