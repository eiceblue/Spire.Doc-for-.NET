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
            // Create a new instance of Document
            Document document = new Document();

            // Load the Word document from the specified file that may contain VBA macros
            document.LoadFromFile(@"../../../../../../Data/Macros.docm", FileFormat.Docm);

            // Save the document to a new file with the specified name and format (Docm for macro-enabled document)
            document.SaveToFile("Sample.docm", FileFormat.Docm);

            // Dispose of the document object when finished using it
            document.Dispose();

            //Launching the Word file.
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
