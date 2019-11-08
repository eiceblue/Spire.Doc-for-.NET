using System;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RetrieveStyle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load a template document 
            Document doc = new Document(@"..\..\..\..\..\..\Data\Styles.docx");

            //Traverse all paragraphs in the document and get their style names through StyleName property
            string styleName = null;
            foreach (Section section in doc.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                     styleName += paragraph.StyleName + "\r\n";
                }
            }

            //Save and launch document
            string output = "RetrieveStyle.txt";
            File.WriteAllText(output, styleName.ToString());
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
