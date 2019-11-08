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

namespace ToXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document.
            Document document = new Document();

            document.LoadFromFile(@"..\..\..\..\..\..\Data\Summary_of_Science.doc");
            //Save the document to a xml file.
            document.SaveToFile("Sample.xml", FileFormat.Xml);

            //Launch the file.
            XMLViewer("Sample.xml");
        }   
        private void XMLViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

    }
}
