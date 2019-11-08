using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemovePageBreaks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create Word document.
            Document document = new Document();

            //Load the file from disk.
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_4.docx");

            //Traverse every paragraph of the first section of the document.
            for (int j = 0; j < document.Sections[0].Paragraphs.Count; j++)
            {
                Paragraph p = document.Sections[0].Paragraphs[j];

                //Traverse every child object of a paragraph.
                for (int i = 0; i < p.ChildObjects.Count; i++)
                {
                    DocumentObject obj = p.ChildObjects[i];

                    //Find the page break object.
                    if (obj.DocumentObjectType == DocumentObjectType.Break)
                    {
                        Break b = obj as Break;

                        //Remove the page break object from paragraph.
                        p.ChildObjects.Remove(b);
                    }
                }
            }

            String result = "Result-RemovePageBreaks.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx2013);

            //Launch the MS Word file.
            WordDocViewer(result);
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
