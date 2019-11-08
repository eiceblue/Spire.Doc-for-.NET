using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CopyContentToAnotherDoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Initialize a new object of Document class and load the source document.
            Document sourceDoc = new Document();
            sourceDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_1.docx");

            //Initialize another object to load target document.
            Document destinationDoc = new Document();
            destinationDoc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Target.docx");

            //Copy content from source file and insert them to the target file.
            foreach (Section sec in sourceDoc.Sections)
            {
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    destinationDoc.Sections[0].Body.ChildObjects.Add(obj.Clone());
                }
            }

            String result = "Result-CopyContentToAnotherWord.docx";

            //Save to file.
            destinationDoc.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
            WordDocViewer(result);
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
