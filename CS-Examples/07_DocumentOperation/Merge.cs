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
                //Create word document
                Document document = new Document();
                document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Summary_of_Science.doc", FileFormat.Doc);

                Document documentMerge = new Document();
                documentMerge.LoadFromFile(@"..\..\..\..\..\..\..\Data\Bookmark.docx", FileFormat.Docx);

                foreach( Section sec in documentMerge.Sections)
                {
                    document.Sections.Add(sec.Clone());
                }

                //Save as docx file.
                document.SaveToFile("Sample.docx", FileFormat.Docx);

                //Launching the MS Word file.
                WordDocViewer("Sample.docx");


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
