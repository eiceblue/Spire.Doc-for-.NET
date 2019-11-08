using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CopyHeaderAndFooter
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

            //Get the header section from the source document
            HeaderFooter header = doc1.Sections[0].HeadersFooters.Header;

            //Load the destination file
            input = @"..\..\..\..\..\..\Data\Template.docx";
            Document doc2 = new Document();
            doc2.LoadFromFile(input);

            //Copy each object in the header of source file to destination file
            foreach (Section section in doc2.Sections)
            {
                foreach (DocumentObject obj in header.ChildObjects)
                {
                    section.HeadersFooters.Header.ChildObjects.Add(obj.Clone());
                }
            }

            //Save and launch document
            string output = "CopyHeaderAndFooter.docx";
            doc2.SaveToFile(output, FileFormat.Docx);
            Viewer(output);
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
