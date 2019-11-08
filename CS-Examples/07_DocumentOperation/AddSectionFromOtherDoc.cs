using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Text;

namespace AddSectionFromOtherDoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Open a Word document as target document
            Document TarDoc = new Document(@"..\..\..\..\..\..\..\Data\SampleB_1.docx");

            //Open a Word document as source document
            Document SouDoc = new Document( @"..\..\..\..\..\..\..\Data\Sample_two sections.docx");

            //Get the second section from source document
            Section Ssection = SouDoc.Sections[1];
           
            //Add the section in target document
            TarDoc.Sections.Add(Ssection.Clone());

            String result = "result.docx";

            //Save to file
            TarDoc.SaveToFile(result, FileFormat.Docx);

            //Launch result file
            WordDocViewer(result);

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
