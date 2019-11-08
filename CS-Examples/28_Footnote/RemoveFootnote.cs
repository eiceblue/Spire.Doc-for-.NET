using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace RemoveFootnote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document document = new Document();
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Footnote.docx");
            Section section = document.Sections[0];

            //traverse paragraphs in the section and find the footnote
            foreach (Paragraph para in section.Paragraphs)
            {
                int index = -1;
                for (int i = 0, cnt = para.ChildObjects.Count; i < cnt; i++)
                {
                    ParagraphBase pBase = para.ChildObjects[i] as ParagraphBase;
                    if (pBase is Footnote)
                    {
                        index = i;
                        break;
                    }
                }

                if (index > -1)
                    //remove the footnote
                    para.ChildObjects.RemoveAt(index);
            }

            document.SaveToFile("RemoveFootnote.docx", FileFormat.Docx);

            //view the Word file.
            WordDocViewer("RemoveFootnote.docx");
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
