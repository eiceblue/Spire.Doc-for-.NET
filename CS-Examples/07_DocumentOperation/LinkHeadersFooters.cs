using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace LinkHeadersFooters
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
            Document srcDoc = new Document();
            srcDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N1.docx");

            //Load the destination file
            Document dstDoc = new Document();
            dstDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N2.docx");

            //Link the headers and footers in the source file
            srcDoc.Sections[0].HeadersFooters.Header.LinkToPrevious = true;
            srcDoc.Sections[0].HeadersFooters.Footer.LinkToPrevious = true;

            //Clone the sections of source to destination
            foreach (Section section in srcDoc.Sections)
            {
                dstDoc.Sections.Add(section.Clone());
            }
            
            //Save the document
            string output="LinkHeadersFooters_out.docx";
            dstDoc.SaveToFile(output, FileFormat.Docx2013);

            //Launching the document
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
