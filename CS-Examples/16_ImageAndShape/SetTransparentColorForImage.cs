using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace SetTransparentColorForImage
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

            //Get the first paragraph in the first section
            Paragraph paragraph = doc.Sections[0].Paragraphs[0];

            //Set the blue color of the image(s) in the paragraph to transperant
            foreach (DocumentObject obj in paragraph.ChildObjects)
            {
                if (obj is DocPicture)
                {
                    DocPicture picture = obj as DocPicture;
                    picture.TransparentColor = Color.Blue;
                }
            }

            //Save and launch document
            string output = "SetTransparentColorForImage.docx";
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
