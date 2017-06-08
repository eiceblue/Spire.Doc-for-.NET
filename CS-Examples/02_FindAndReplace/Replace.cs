using System;
using System.Windows.Forms;
using Spire.Doc;

namespace Replace
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

            //load a document
            document.LoadFromFile(@"..\..\..\..\..\..\Data\FindAndReplace.doc");

            //Replace text
            document.Replace(this.textBox1.Text, this.textBox2.Text,true,true);

            //Save doc file.
            document.SaveToFile("Sample.doc", FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");
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
