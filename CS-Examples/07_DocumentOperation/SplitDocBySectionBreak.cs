using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace SplitDocBySectionBreak
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
            document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_4.docx");

            //Define another new word document object.
            Document newWord;

            //Split a Word document into multiple documents by section break.
            for (int i = 0; i < document.Sections.Count; i++)
            {
                String result = String.Format("Result-SplitWordFileBySectionBreak_{0}.docx", i);
                newWord = new Document();
                newWord.Sections.Add(document.Sections[i].Clone());

                //Save to file.
                newWord.SaveToFile(result);

                //Launch the MS Word file.
                WordDocViewer(result);
            }
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
