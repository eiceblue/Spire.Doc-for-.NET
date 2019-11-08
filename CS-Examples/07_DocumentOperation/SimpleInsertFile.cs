using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using System.IO;

namespace SimpleInsertFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load the Word document
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_N5.docx");

            //Insert document from file
            doc.InsertTextFromFile(@"..\..\..\..\..\..\..\Data\Template_N3.docx", FileFormat.Auto);

            //Save the document
            string output="SimpleInsertFile_out.docx";
            doc.SaveToFile(output,FileFormat.Docx2013);

            //Launch the document
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
