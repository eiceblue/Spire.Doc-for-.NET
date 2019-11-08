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
            //Load the source file
            string input = @"..\..\..\..\..\..\Data\HeaderAndFooter.docx";
            Document doc1 = new Document();
            doc1.LoadFromFile(input);

            //Get the header from the first section
            HeaderFooter header = doc1.Sections[0].HeadersFooters.Header;

            //Load the destination file
            input = @"..\..\..\..\..\..\Data\MultiplePages.docx";
            Document doc2 = new Document();
            doc2.LoadFromFile(input);

            //Get the first page header of the destination document
            HeaderFooter firstPageHeader = doc2.Sections[0].HeadersFooters.FirstPageHeader;

            //Specify that the current section has a different header/footer for the first page
            foreach (Section section in doc2.Sections)
            {
                section.PageSetup.DifferentFirstPageHeaderFooter = true;
            }

            //Removes all child objects in firstPageHeader
            firstPageHeader.Paragraphs.Clear();

            //Add all child objects of the header to firstPageHeader
            foreach (DocumentObject obj in header.ChildObjects)
            {
                firstPageHeader.ChildObjects.Add(obj.Clone());
            }

            //Save and launch the file
            string resultfile = "AddHeaderOnlyFirstPage.docx";
            doc2.SaveToFile(resultfile, FileFormat.Docx);
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
