using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace UpdateImage
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

            //Get all pictures in the Word document
            List<DocumentObject> pictures = new List<DocumentObject>();
            foreach (Section sec in doc.Sections)
            {
                foreach (Paragraph para in sec.Paragraphs)
                { 
                    foreach (DocumentObject docObj in para.ChildObjects)
                    {
                        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            pictures.Add(docObj);
                        }
                    }
                }
            }

            //Replace the first picture with a new image file
            DocPicture picture = pictures[0] as DocPicture;
            picture.LoadImage(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));

            //Save and launch document
            string output = "ReplaceWithNewImage.docx";
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
