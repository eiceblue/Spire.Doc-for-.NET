using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace WordToWordXML
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");
          
            String result1 = "Result-WordToWordML.xml";          
            String result2 = "Result-WordToWordXML.xml";

            //For word 2003:
            document.SaveToFile(result1, FileFormat.WordML);

            //For word 2007:
            document.SaveToFile(result2, FileFormat.WordXml);

            //Launch the files.
            WordDocViewer(result1);
            WordDocViewer(result2);
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
