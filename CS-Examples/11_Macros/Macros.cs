using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml.XPath;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Macros
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document document = new Document();

            //Loading documetn with macros.
            document.LoadFromFile(@"../../../../../../Data/Macros.docm", FileFormat.Docm);

            //Removes the macros from the document.
            document.ClearMacros();

            //Save docm file.
            document.SaveToFile("Sample.docm", FileFormat.Docm);

            //Launching the MS Word file.
            WordDocViewer("Sample.docm");
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
