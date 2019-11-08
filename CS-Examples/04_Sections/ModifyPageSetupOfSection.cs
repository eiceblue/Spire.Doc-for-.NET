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
            //Load Word from disk
            Document doc = new Document();
            doc.LoadFromFile(@"../../../../../../Data/Template_N2.docx");

            //Loop through all sections
            foreach (Section section in doc.Sections)
            {
                //Modify the margins
                section.PageSetup.Margins = new MarginsF(100, 80, 100, 80);
                //Modify the page size
                section.PageSetup.PageSize = PageSize.Letter;
            }

            // Or only modify one section
            // For example, modify the page setup of the first section
            //Section section0 = doc.Sections[0];
            //section0.PageSetup.Margins = new MarginsF(100, 80, 100, 80);
            //section0.PageSetup.FooterDistance = 35.4f;
            //section0.PageSetup.HeaderDistance = 34.4f;

            //Save the Word file
            string output = "ModifyPageSetupOfAllSections_out.docx";
            doc.SaveToFile(output,FileFormat.Docx2013);

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
