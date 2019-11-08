using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;

namespace RemoveContentControls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Load document from disk
            Document doc = new Document();
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveContentControls.docx");

            //Loop through sections
            for (int s = 0; s < doc.Sections.Count; s++)
            {
                Section section = doc.Sections[s];
                for (int i = 0; i < section.Body.ChildObjects.Count; i++)
                {
                    //Loop through contents in paragraph
                    if (section.Body.ChildObjects[i] is Paragraph)
                    {
                        Paragraph para = section.Body.ChildObjects[i] as Paragraph;
                        for (int j = 0; j < para.ChildObjects.Count; j++)
                        {
                            //Find the StructureDocumentTagInline
                            if (para.ChildObjects[j] is StructureDocumentTagInline)
                            {
                                StructureDocumentTagInline sdt = para.ChildObjects[j] as StructureDocumentTagInline;
                                //Remove the content control from paragraph
                                para.ChildObjects.Remove(sdt);
                                j--;
                            }
                        }
                    }
                    if (section.Body.ChildObjects[i] is StructureDocumentTag)
                    {
                        StructureDocumentTag sdt = section.Body.ChildObjects[i] as StructureDocumentTag;
                        section.Body.ChildObjects.Remove(sdt);
                        i--;
                    }
                }
            }

            //Save the Word document
            string output = "RemoveContentControls_out.docx";
            doc.SaveToFile(output, FileFormat.Docx2013);

            //Launch the file
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
