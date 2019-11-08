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

namespace AddAndDeleteSections
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\SectionTemplate.docx");

            AddSection(doc);
            DeleteSection(doc);

            string output="AddAndDeleteSections_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

            FileViewer(output);
        }
        private void AddSection(Document doc)
        {
            //Add a section
            doc.AddSection();
        }
        private void DeleteSection(Document doc)
        {
            //Delete the last section
            doc.Sections.RemoveAt(doc.Sections.Count - 1);
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
