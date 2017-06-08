using System;
using System.Windows.Forms;
using Spire.Doc;

namespace Decrypt
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
            if (!string.IsNullOrEmpty(fileName))
            {
                //Create word document
                Document document = new Document();
                document.LoadFromFile(fileName,FileFormat.Doc,this.textBox1.Text);

                //Save doc file.
                document.SaveToFile("Sample.doc", FileFormat.Doc);

                //Launching the MS Word file.
                WordDocViewer("Sample.doc");
            }


        }

        private string OpenFile()
        {
            openFileDialog1.InitialDirectory
                = new System.IO.DirectoryInfo(@"..\..\..\..\..\..\Data").FullName;
            openFileDialog1.FileName = "Protect_Decrypt.doc";
            openFileDialog1.Filter = "Word Document (*.doc)|*.doc";
            openFileDialog1.Title = "Choose a document to Decrypt";

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
