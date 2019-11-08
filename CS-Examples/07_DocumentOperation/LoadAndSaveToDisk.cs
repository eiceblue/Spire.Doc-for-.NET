using System;
using System.Windows.Forms;
using System.Text;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Formatting;
using System.IO;
using Spire.Doc.Fields;
namespace LoadAndSaveToDisk
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String input = @"..\..\..\..\..\..\Data\Template.docx";
            //Create a new document
            Document doc = new Document();
            // Load the document from the absolute/relative path on disk.
            doc.LoadFromFile(input);

            String result = "LoadAndSaveToDisk_out.docx";
            // Save the document to disk
            doc.SaveToFile(result,FileFormat.Docx);
            WordViewer(result);
        }
        private void WordViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
