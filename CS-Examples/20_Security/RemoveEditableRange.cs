using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace RemoveEditableRange
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create a new document
            Document document = new Document();
            //Load file from disk
            document.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveEditableRange.docx");
            //Find "PermissionStart" and "PermissionEnd" tags and remove them
            foreach(Section section in document.Sections)
            {
                foreach(Paragraph paragraph in section.Body.Paragraphs)
                {
                    for(int i=0;i<paragraph.ChildObjects.Count;)
                    {
                        DocumentObject obj = paragraph.ChildObjects[i];
                        if(obj is PermissionStart||obj is PermissionEnd)
                        {
                            paragraph.ChildObjects.Remove(obj);
                        }
                        else
                        {
                            i++;
                        }
                    }
                }
            }
            //Save the document
            string output = "RemoveEditableRange_output.docx";
            document.SaveToFile(output,FileFormat.Docx);
            WordDocViewer(output);
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
