using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace KeepSameFormat
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the source document from disk
            Document srcDoc = new Document();
            srcDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N2.docx");

            //Load the destination document from disk
            Document destDoc = new Document();
            destDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N3.docx");

            //Keep same format of source document
            srcDoc.KeepSameFormat = true;

            //Copy the sections of source document to destination document
            foreach (Section section in srcDoc.Sections)
            {
                destDoc.Sections.Add(section.Clone());
            }

            //Save the Word document
            string output="KeepSameFormating_out.docx";
            destDoc.SaveToFile(output, FileFormat.Docx2013);

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
