using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ChangeCase
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a new document and load from file
            string input = @"..\..\..\..\..\..\Data\Text1.docx"; ;
            Document doc = new Document();
            doc.LoadFromFile(input);
            TextRange textRange;
            //Get the first paragraph and set its CharacterFormat to AllCaps
            Paragraph para1 = doc.Sections[0].Paragraphs[1];

            foreach (DocumentObject obj in para1.ChildObjects)
            {
                if (obj is TextRange)
                {
                    textRange = obj as TextRange;
                    textRange.CharacterFormat.AllCaps = true;
                }
            }
     
            //Get the third paragraph and set its CharacterFormat to IsSmallCaps
            Paragraph para2 = doc.Sections[0].Paragraphs[3];
            foreach (DocumentObject obj in para2.ChildObjects)
            {
                if (obj is TextRange)
                {
                    textRange = obj as TextRange;
                    textRange.CharacterFormat.IsSmallCaps = true;
                }
            }
         

            //Save and launch the document
            string output = "ChangeCase.docx";
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
