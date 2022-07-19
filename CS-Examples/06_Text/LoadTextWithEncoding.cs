using System;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace LoadTextWithEncoding
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string inputFile = "../../../../../../Data/Sample_UTF-7.txt";
            Document document = new Document();
            document.LoadText(inputFile, Encoding.UTF7);
            string resultFile = "LoadTextWithEncoding_out.docx";
            document.SaveToFile(resultFile, FileFormat.Docx);
            WordDocViewer(resultFile);

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
