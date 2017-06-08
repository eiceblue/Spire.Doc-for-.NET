using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace FontAndColor
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
            String text
                = "This paragraph is demo of text font and color. "
                + "The font name of this paragraph is Tahoma. "
                + "The font size of this paragraph is 20. "
                + "The under line style of this paragraph is DotDot. "
                + "The color of this paragraph is Blue. ";
            TextRange txtRange = paragraph.AppendText(text);

            //Font name
            txtRange.CharacterFormat.FontName = "Tahoma";

            //Font size
            txtRange.CharacterFormat.FontSize = 20;

            //Underline
            txtRange.CharacterFormat.UnderlineStyle = UnderlineStyle.DotDot;

            //Change text color
            txtRange.CharacterFormat.TextColor = Color.Blue;

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
