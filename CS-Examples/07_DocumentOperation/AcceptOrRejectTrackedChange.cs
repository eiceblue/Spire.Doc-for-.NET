using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace AcceptOrRejectTrackedChange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\AcceptOrRejectTrackedChanges.docx");

            //Get the first section and the paragraph we want to accept/reject the changes.
            Section sec = document.Sections[0];
            Paragraph para = sec.Paragraphs[0];

            //Accept the changes or reject the changes.
            para.Document.AcceptChanges();
            //para.Document.RejectChanges();

            String result = "Result-AcceptOrRejectTrackedChanges.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

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
