using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SplitText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document and load from file
            string input = @"..\..\..\..\..\..\Data\Sample.docx"; ;
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Add a column to the first section and set width and spacing
            doc.Sections[0].AddColumn(100f, 20f);
            //Add a line between the two columns
            doc.Sections[0].PageSetup.ColumnsLineBetween = true;

            //Save and launch the document
            string output = "SplitText.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);
            Viewer(output);
        }
        private void Viewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
