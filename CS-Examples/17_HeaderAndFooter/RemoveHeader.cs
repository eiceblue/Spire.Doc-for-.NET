using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveHeader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the document
            string input = @"..\..\..\..\..\..\Data\HeaderAndFooter.docx";            
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first section of the document
            Section section = doc.Sections[0];

            //Traverse the word document and clear all headers in different type
            foreach (Paragraph para in section.Paragraphs)
            {
                foreach (DocumentObject obj in para.ChildObjects)
                {
                    //Clear footer in the first page
                    HeaderFooter header;
                    header = section.HeadersFooters[HeaderFooterType.HeaderFirstPage];
                    if (header != null)
                        header.ChildObjects.Clear();
                    //Clear footer in the odd page
                    header = section.HeadersFooters[HeaderFooterType.HeaderOdd];
                    if (header != null)
                        header.ChildObjects.Clear();
                    //Clear footer in the even page
                    header = section.HeadersFooters[HeaderFooterType.HeaderEven];
                    if (header != null)
                        header.ChildObjects.Clear();
                }
            }

            //Save and launch document
            string output = "RemoveHeader.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
