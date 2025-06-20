using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Text;

namespace SectionBreakContinuous
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Create a new instance of the Document class and load a document from the specified file path.
			Document doc = new Document(@"..\..\..\..\..\..\..\Data\Sample_two sections.docx");

			// Iterate through each section in the document.
			foreach (Section section in doc.Sections)
			{
				// Set the break code of each section to NoBreak, which means no section break will be inserted.
				section.BreakCode = SectionBreakType.NoBreak;
			}

			// Save the modified document to a file with the specified file name.
			doc.SaveToFile("result.docx");

			// Dispose of the document to release resources.
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
