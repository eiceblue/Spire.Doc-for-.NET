using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Text;

namespace DifferentPageSetup
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Create a new instance of the Document class and load a Word document from a specific file path
			Document doc = new Document(@"..\..\..\..\..\..\Data\DifferentPageSetup.docx");

			// Get the second section of the document
			Section SectionTwo = doc.Sections[1];

			// Set the page orientation of the second section to Landscape
			SectionTwo.PageSetup.Orientation = PageOrientation.Landscape;

			// Uncomment the following line to set a custom page size for the second section
			// SectionTwo.PageSetup.PageSize = new SizeF(800, 800);

			// Save the modified document to a file named "result.docx"
			doc.SaveToFile("result.docx");

			// Release the resources used by the document object
			doc.Dispose();

            //Launch result file
            WordDocViewer("result.docx");

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
