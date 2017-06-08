using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Text
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
            paragraph.AppendText("The various ways to format paragraph text in Microsoft Word:");

            paragraph.ApplyStyle(BuiltinStyle.Heading1);

            //Append alignment text
            AppendAligmentText(section);

            //Append indentation text
            AppendIndentationText(section);

            AppendBulletedList(section);

            //Save doc file.
            document.SaveToFile("Sample.doc",FileFormat.Doc);

            //Launching the MS Word file.
            WordDocViewer("Sample.doc");


        }

        private void AppendAligmentText(Section section)
        {
            Paragraph paragraph = null;

            paragraph = section.AddParagraph();

            //Append Text
            paragraph.AppendText("Horizontal Aligenment");

            paragraph.ApplyStyle(BuiltinStyle.Heading3);

            foreach (Spire.Doc.Documents.HorizontalAlignment align in Enum.GetValues(typeof(Spire.Doc.Documents.HorizontalAlignment)))
            {
                Paragraph paramgraph = section.AddParagraph();
                paramgraph.AppendText("This text is " + align.ToString());
                paramgraph.Format.HorizontalAlignment = align;
            }
        }

        private void AppendIndentationText(Section section)
        {
            Paragraph paragraph = null;

            paragraph = section.AddParagraph();

            //Append Text
            paragraph.AppendText("Indentation");

            paragraph.ApplyStyle(BuiltinStyle.Heading3);

            paragraph = section.AddParagraph();
            paragraph.AppendText("Indentation is the spacing between text and margins. Word allows you to set left and right margins, as well as indentations for the first line of a paragraph and hanging indents");
            paragraph.Format.FirstLineIndent = 15;
        }

        private void AppendBulletedList(Section section)
        {
            Paragraph paragraph = null;

            paragraph = section.AddParagraph();
            

            //Append Text
            paragraph.AppendText("Bulleted List");

            paragraph.ApplyStyle(BuiltinStyle.Heading3);

            paragraph = section.AddParagraph();
            for (int i = 0; i < 5; i++)
            {
                paragraph = section.AddParagraph();
                paragraph.AppendText("Item" + i.ToString());

                if (i == 0)
                {
                    paragraph.ListFormat.ApplyBulletStyle();
                }
                else
                {
                    paragraph.ListFormat.ContinueListNumbering();
                }

                paragraph.ListFormat.ListLevelNumber = 1;
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
