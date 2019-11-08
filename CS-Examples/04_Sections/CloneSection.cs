using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace CloneSection
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load source file
            Document srcDoc = new Document();
            srcDoc.LoadFromFile(@"..\..\..\..\..\..\Data\SectionTemplate.docx");
            
            //Create destination file
            Document desDoc = new Document();
            
            Section cloneSection=null;
            foreach (Section section in srcDoc.Sections)
            {
                //Clone section
                cloneSection = section.Clone();
                //Add the cloneSection in destination file
                desDoc.Sections.Add(cloneSection);
            }

            //Save the Word
            string output="CloneSection_out.docx";
            desDoc.SaveToFile(output, FileFormat.Docx2013);

            //Launch Word file
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