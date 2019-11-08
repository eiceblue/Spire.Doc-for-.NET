using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace RemoveEmptyLines
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_3.docx");

            //Traverse every section on the word document and remove the null and empty paragraphs.
            foreach (Section section in document.Sections)
            {
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    if (section.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        if (String.IsNullOrEmpty((section.Body.ChildObjects[i] as Paragraph).Text.Trim()))
                        {
                            section.Body.ChildObjects.Remove(section.Body.ChildObjects[i]);
                            i--;
                        }
                    }

                }
            }

            String result = "Result-RemoveEmptyLines.docx";

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
