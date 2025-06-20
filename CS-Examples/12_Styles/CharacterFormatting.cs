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
            //Create a Word document.
			Document document = new Document();

			//Add a section
			Section sec = document.AddSection();

			//Add a paragraph
			Paragraph titleParagraph = sec.AddParagraph();

			//Append text
			titleParagraph.AppendText("Font Styles and Effects ");

			//Apply the builtin style
			titleParagraph.ApplyStyle(BuiltinStyle.Title);

			//Add a new paragraph
			Paragraph paragraph = sec.AddParagraph();

			//Append text
			TextRange tr = paragraph.AppendText("Strikethough Text");

			//Set strikeout style
			tr.CharacterFormat.IsStrikeout = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Shadow Text");

			//Set shadow property of text
			tr.CharacterFormat.IsShadow = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Small caps Text");

			//Set IsSmallCaps property of text
			tr.CharacterFormat.IsSmallCaps = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Double Strikethough Text");

			//Set DoubleStrike property of text
			tr.CharacterFormat.DoubleStrike = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Outline Text");
			tr.CharacterFormat.IsOutLine = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("AllCaps Text");
			tr.CharacterFormat.AllCaps = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			paragraph.AppendText("Text");
			tr = paragraph.AppendText("SubScript");

			//Apply CharacterFormat
			tr.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

			//Append text
			tr = paragraph.AppendText("And");
			tr = paragraph.AppendText("SuperScript");

			//Apply CharacterFormat
			tr.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Emboss Text");

			//Apply CharacterFormat
			tr.CharacterFormat.Emboss = true;
			tr.CharacterFormat.TextColor = Color.White;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			paragraph.AppendText("Hidden:");
			tr = paragraph.AppendText("Hidden Text");

			//Apply CharacterFormat
			tr.CharacterFormat.Hidden = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Engrave Text");

			//Apply CharacterFormat
			tr.CharacterFormat.Engrave = true;
			tr.CharacterFormat.TextColor = Color.White;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("WesternFonts╓╨╬─╫╓╠х");

			//Apply CharacterFormat
			tr.CharacterFormat.FontNameAscii = "Calibri";
			tr.CharacterFormat.FontNameNonFarEast = "Calibri";
			tr.CharacterFormat.FontNameFarEast = "Simsun";

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Font Size");

			//Apply CharacterFormat
			tr.CharacterFormat.FontSize = 20;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Font Color");

			//Apply CharacterFormat
			tr.CharacterFormat.TextColor = Color.Red;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Bold Italic Text");

			//Apply CharacterFormat
			tr.CharacterFormat.Bold = true;
			tr.CharacterFormat.Italic = true;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Underline Style");

			//Apply CharacterFormat
			tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Highlight Text");

			//Apply CharacterFormat
			tr.CharacterFormat.HighlightColor = Color.Yellow;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Text has shading");

			//Apply CharacterFormat
			tr.CharacterFormat.TextBackgroundColor = Color.Green;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Border Around Text");

			//Apply CharacterFormat
			tr.CharacterFormat.Border.BorderType = Spire.Doc.Documents.BorderStyle.Single;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Text Scale");

			//Apply CharacterFormat
			tr.CharacterFormat.TextScale = 150;

			//Append a line break
			paragraph.AppendBreak(BreakType.LineBreak);

			//Append text
			tr = paragraph.AppendText("Character Spacing is 2 point");

			//Apply CharacterFormat
			tr.CharacterFormat.CharacterSpacing = 2;

			string filePath = "CharaterFormatting.docx";
			document.SaveToFile(filePath, FileFormat.Docx);

			//Dispose the document
			document.Dispose();
            WordDocViewer(filePath);
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
