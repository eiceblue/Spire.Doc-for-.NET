using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace InsertRtfStringToDoc
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

            //Add a new section.
            Section section = document.AddSection();

            //Add a paragraph to the section.
            Paragraph para = section.AddParagraph();

            //Declare a String variable to store the Rtf string.
            String rtfString = @"{\rtf1\ansi\deff0 {\fonttbl {\f0 hakuyoxingshu7000;}}\f0\fs28 Hello, World}";

            //Append Rtf string to paragraph.
            para.AppendRTF(rtfString);

            String result = "Result-InsertRtfStringToWord.docx";

            //Save to file.
            document.SaveToFile(result, FileFormat.Docx);

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
