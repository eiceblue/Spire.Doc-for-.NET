using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ModifyPageSetupOfSection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new instance of the Document class.
            Document doc = new Document();

            // Load a Word document from a specified file path using the LoadFromFile method.
            doc.LoadFromFile(@"../../../../../../Data/Template_N2.docx");

            // Iterate through each section in the document.
            foreach (Section section in doc.Sections)
            {
                // Set the page margins of the current section using the MarginsF class.
                section.PageSetup.Margins = new MarginsF(100, 80, 100, 80);

                // Set the page size of the current section to Letter size.
                section.PageSetup.PageSize = PageSize.Letter;
            }
			
			// Or only modify one section
            // For example, modify the page setup of the first section
            //Section section0 = doc.Sections[0];
            //section0.PageSetup.Margins = new MarginsF(100, 80, 100, 80);
            //section0.PageSetup.FooterDistance = 35.4f;
            //section0.PageSetup.HeaderDistance = 34.4f;

            // Specify the output file name for the modified page setup document.
            string output = "ModifyPageSetupOfAllSections_out.docx";

            // Save the document to a file with the specified output file name and Docx2013 format.
            doc.SaveToFile(output, FileFormat.Docx2013);

            // Release system resources used by the document.
            doc.Dispose();


            //Launch the file
            WordDocViewer(output);
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
