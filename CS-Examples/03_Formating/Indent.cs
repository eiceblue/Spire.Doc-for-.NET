using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Indent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Open a blank word document as template
            Document document = new Document(@"..\..\..\..\..\..\Data\Blank.doc");

            //Get the first secition
            Section section = document.Sections[0];

            //Create a new paragraph or get the first paragraph
            Paragraph paragraph
                = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

            //Append Text
            paragraph.AppendText("Using items list to show Indent demo.");
            paragraph.ApplyStyle(BuiltinStyle.Heading3);

            paragraph = section.AddParagraph();
            for (int i = 0; i < 10; i++)
            {
                paragraph = section.AddParagraph();
                paragraph.AppendText(String.Format("Indent Demo Node{0}", i));

                
                if(i == 0)
                {
                    paragraph.ListFormat.ApplyBulletStyle();
                }
                else
                {
                    paragraph.ListFormat.ContinueListNumbering();
                }

                paragraph.ListFormat.CurrentListLevel.NumberPosition = -10;
            }

            //Save doc file.
            document.SaveToFile("Sample.doc",FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");
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
