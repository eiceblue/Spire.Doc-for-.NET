using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Styles
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
            paragraph.AppendText("Builtin Style:");

            foreach (BuiltinStyle builtinStyle in Enum.GetValues(typeof(BuiltinStyle)))
            {
                paragraph = section.AddParagraph();
                //Append Text
                paragraph.AppendText(builtinStyle.ToString());
                //Apply Style
                paragraph.ApplyStyle(builtinStyle);
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
