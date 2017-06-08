using System;
using System.Windows.Forms;
using Spire.Doc;

namespace Merge
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = OpenFile();
            string fileMerge = OpenFile();
            if ((!string.IsNullOrEmpty(fileName)) && (!string.IsNullOrEmpty(fileMerge)))
            {
                //Create word document
                Document document = new Document();
                document.LoadFromFile(fileName,FileFormat.Doc);

                Document documentMerge = new Document();
                documentMerge.LoadFromFile(fileMerge, FileFormat.Doc);

                foreach( Section sec in documentMerge.Sections)
                {
                    document.Sections.Add(sec.Clone());
                }

                //Save doc file.
                document.SaveToFile("Sample.doc", FileFormat.Doc);

                //Launching the MS Word file.
                WordDocViewer("Sample.doc");
            }


        }

        private string OpenFile()
        {
            openFileDialog1.Filter = "Word Document (*.doc)|*.doc";
            openFileDialog1.Title = "Choose a document to merage";

            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog1.FileName;
            }

            return string.Empty;
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
