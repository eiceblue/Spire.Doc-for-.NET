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
            // Create a new Document object.
            Document document = new Document();

            // Add a new Section to the document.
            Section section = document.AddSection();

            // Add a new Paragraph to the Section.
            Paragraph para = section.AddParagraph();

            // Define an RTF string containing formatted text.
            String rtfString = @"{\rtf1\ansi\deff0 {\fonttbl {\f0 hakuyoxingshu7000;}}\f0\fs28 Hello, World}";

            // Append the RTF string to the Paragraph, preserving the formatting.
            para.AppendRTF(rtfString);

            // Specify the file name for the resulting Word document.
            String result = "Result-InsertRtfStringToWord.docx";

            // Save the Document object to a file in Docx format.
            document.SaveToFile(result, FileFormat.Docx);

            // Dispose the Document object.
            document.Dispose();

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
