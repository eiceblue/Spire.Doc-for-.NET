using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace ChangeTOCTabStyle
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
            doc.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Toc.docx");

            //Loop through sections
            foreach (Section section in doc.Sections)
            {
                //Loop through content of section
                foreach (DocumentObject obj in section.Body.ChildObjects)
                {
                    //Find the structure document tag
                    if (obj is StructureDocumentTag)
                    {
                        StructureDocumentTag tag = obj as StructureDocumentTag;
                        //Find the paragraph where the TOC1 locates
                        foreach (DocumentObject cObj in tag.ChildObjects)
                        {
                            if (cObj is Paragraph)
                            {
                                Paragraph para = cObj as Paragraph;
                                if (para.StyleName == "TOC2")
                                {
                                    //Set the tab style of paragraph
                                    foreach (Tab tab in para.Format.Tabs)
                                    {
                                        tab.Position = tab.Position + 20;
                                        tab.TabLeader = TabLeader.NoLeader;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //Save the Word file
            string output = "ChangeTOCTabStyle_out.docx";
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
