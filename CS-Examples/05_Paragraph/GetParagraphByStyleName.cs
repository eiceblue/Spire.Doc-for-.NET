using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;

namespace GetParagraphByStyleName
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
            document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_DocX_3.docx");

            StringBuilder content = new StringBuilder();
            content.AppendLine("Get paragraphs by style name \"Heading1\": ");

            //Get paragraphs by style name.
            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    if (paragraph.StyleName == "Heading1")
                    {
                        content.AppendLine(paragraph.Text);
                    }
                }
            }

            String result = "Result-GetParagraphsByStyleName.txt";

            //Save to file.
            File.WriteAllText(result,content.ToString());

            //Launch the file.
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
