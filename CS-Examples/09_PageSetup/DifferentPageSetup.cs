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

            //Open a Word document
            Document doc = new Document(@"..\..\..\..\..\..\Data\DifferentPageSetup.docx");

            //Get the second section 
            Section SectionTwo = doc.Sections[1];

            //Set the orientation
            SectionTwo.PageSetup.Orientation = PageOrientation.Landscape;

            //Set page size
            //SectionTwo.PageSetup.PageSize = new SizeF(800, 800);

            doc.SaveToFile("result.docx");

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
