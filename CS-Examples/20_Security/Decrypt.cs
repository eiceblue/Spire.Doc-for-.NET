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
                //Create word document
                Document document = new Document();
                document.LoadFromFile(@"..\..\..\..\..\..\Data\TemplateWithPassword.docx", FileFormat.Docx, "E-iceblue");

                //Save as doc file.
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
