using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CloneSectionContent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the Word document from disk
            Document doc=new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\SectionTemplate.docx");

            //Get the first section
            Section sec1 = doc.Sections[0];
            //Get the second section
            Section sec2 = doc.Sections[1];

            //Loop through the contents of sec1
            foreach (DocumentObject obj in sec1.Body.ChildObjects)
            {
                //Clone the contents to sec2
                sec2.Body.ChildObjects.Add(obj.Clone());
            }

            //Save the Word document
            string output = "CloneSectionContent_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

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
