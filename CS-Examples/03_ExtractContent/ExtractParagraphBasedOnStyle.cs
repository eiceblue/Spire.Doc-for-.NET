using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace ExtractParagraphBasedOnStyle
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
            String styleName1 = "Heading1";
            StringBuilder style1Text = new StringBuilder();
            //Load file from disk
            document.LoadFromFile(@"..\..\..\..\..\..\Data\ExtractParagraphBasedOnStyle.docx");
            style1Text.AppendLine("The following is the content of the paragraph with the style name " + styleName1 + ": ");
            //Extrct paragraph based on style
            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    if (paragraph.StyleName != null && paragraph.StyleName.Equals(styleName1))
                    {
                        style1Text.AppendLine(paragraph.Text);
                    }
                }
            }
            //Save the content with style "Heading 1"
            string output1 = "ExtractParagraphBasedOnStyle_style1.txt";
            File.WriteAllText(output1, style1Text.ToString());
            WordDocViewer(output1);
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
