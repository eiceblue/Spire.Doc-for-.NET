using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ToMhtml
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

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ToMhtml.docx");

            //Save to RTF file.
            document.SaveToFile("ToMhtml-out.mhtml", FileFormat.Mhtml);

            //Dispose the document
            document.Dispose();

            //Launching the MS Word file.
            WordDocViewer("ToMhtml-out.mhtml");
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
