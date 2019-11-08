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

namespace HideParagraph
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

            //Get the first section and the first paragraph from the word document.
            Section sec = document.Sections[0];
            Paragraph para = sec.Paragraphs[0];
         
            //Loop through the textranges and set CharacterFormat.Hidden property as true to hide the texts.
            foreach (DocumentObject obj in para.ChildObjects)
            {
                if (obj is TextRange)
                {
                    TextRange range = obj as TextRange;
                    range.CharacterFormat.Hidden = true;
                }
            }

            String result = "Result-HideWordParagraph.docx";

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
