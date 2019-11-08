using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace RemoveSectionContent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Word file from disk
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_N3.docx");

            //Loop through all sections
            foreach (Section section in doc.Sections)
            {
                //Remove header content
                section.HeadersFooters.Header.ChildObjects.Clear();
                //Remove body content
                section.Body.ChildObjects.Clear();
                //Remove footer content
                section.HeadersFooters.Footer.ChildObjects.Clear();
            }

            //Save the Word file
            string output="RemoveSectionContent_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

            //Launch the file
            FileViewer(output);
        }
        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
