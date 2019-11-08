using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;

namespace SetTextWrap
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

            foreach (Section sec in doc.Sections)
            {
                foreach (Paragraph para in sec.Paragraphs)
                {
                    List<DocumentObject> pictures = new List<DocumentObject>();
                    //Get all pictures in the Word document
                    foreach (DocumentObject docObj in para.ChildObjects)
                    {
                        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            pictures.Add(docObj);
                        }
                    }

                    //Set text wrap styles for each piture
                    foreach (DocumentObject pic in pictures)
                    {
                        DocPicture picture = pic as DocPicture;
                        picture.TextWrappingStyle = TextWrappingStyle.Through;
                        picture.TextWrappingType = TextWrappingType.Both;
                    }
                }
            }

            //Save and launch document
            string output = "SetTextWrap.docx";
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
