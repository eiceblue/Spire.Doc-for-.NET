using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ResetImageSize
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load Document
            string input = @"..\..\..\..\..\..\Data\ImageTemplate.docx";
            Document doc = new Document();
            doc.LoadFromFile(input);

            //Get the first secion
            Section section = doc.Sections[0];
            //Get the first paragraph
            Paragraph paragraph = section.Paragraphs[0];

            //Reset the image size of the first paragraph
            foreach (DocumentObject docObj in paragraph.ChildObjects)
            {
                if (docObj is DocPicture)
                {
                    DocPicture picture = docObj as DocPicture;
                    picture.Width = 50f;
                    picture.Height = 50f;
                }
            }

            //Save and launch document
            string output = "ResetImageSize.docx";
            doc.SaveToFile(output, FileFormat.Docx);
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
