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

            //Open a Word document
            Document doc = new Document(@"..\..\..\..\..\..\..\Data\Sample_two sections.docx");

            foreach (Section section in doc.Sections)
            {
                //Set section break as continuous
                section.BreakCode = SectionBreakType.NoBreak;
            }

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
