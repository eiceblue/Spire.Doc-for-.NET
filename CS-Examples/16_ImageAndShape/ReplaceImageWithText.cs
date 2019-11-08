using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Collections.Generic;
using Spire.Doc.Fields;

namespace ReplaceImageWithText
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

            //Replace all pictures with texts
            int j = 1;
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

                    //Replace pitures with the text "Here was image {image index}"
                    foreach (DocumentObject pic in pictures)
                    {
                        int index = para.ChildObjects.IndexOf(pic);
                        TextRange range = new TextRange(doc);
                        range.Text = string.Format("Here was image {0}", j);
                        para.ChildObjects.Insert(index, range);
                        para.ChildObjects.Remove(pic);
			            j++;
                    }
                }
            }

            //Save and launch document
            string output = "ReplaceWithTexts.docx";
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
