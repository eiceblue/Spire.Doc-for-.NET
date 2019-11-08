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
            //Initialize a document
            Document document = new Document();
            Section sec = document.AddSection();
            Paragraph titleParagraph = sec.AddParagraph();
            titleParagraph.AppendText("Font Styles and Effects ");
            titleParagraph.ApplyStyle(BuiltinStyle.Title);

            Paragraph paragraph = sec.AddParagraph();
            TextRange tr = paragraph.AppendText("Strikethough Text");
            tr.CharacterFormat.IsStrikeout = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Shadow Text");
            tr.CharacterFormat.IsShadow = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Small caps Text");
            tr.CharacterFormat.IsSmallCaps = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Double Strikethough Text");
            tr.CharacterFormat.DoubleStrike = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Outline Text");
            tr.CharacterFormat.IsOutLine = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("AllCaps Text");
            tr.CharacterFormat.AllCaps = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Text");
            tr = paragraph.AppendText("SubScript");
            tr.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

            tr = paragraph.AppendText("And");
            tr = paragraph.AppendText("SuperScript");
            tr.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Emboss Text");
            tr.CharacterFormat.Emboss = true;
            tr.CharacterFormat.TextColor = Color.White;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Hidden:");
            tr = paragraph.AppendText("Hidden Text");
            tr.CharacterFormat.Hidden = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Engrave Text");
            tr.CharacterFormat.Engrave = true;
            tr.CharacterFormat.TextColor = Color.White;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("WesternFontsÖÐÎÄ×ÖÌå");
            tr.CharacterFormat.FontNameAscii = "Calibri";
            tr.CharacterFormat.FontNameNonFarEast = "Calibri";
            tr.CharacterFormat.FontNameFarEast = "Simsun";

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Font Size");
            tr.CharacterFormat.FontSize = 20;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Font Color");
            tr.CharacterFormat.TextColor=Color.Red;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Bold Italic Text");
            tr.CharacterFormat.Bold = true;
            tr.CharacterFormat.Italic = true;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Underline Style");
            tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Highlight Text");
            tr.CharacterFormat.HighlightColor = Color.Yellow;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Text has shading");
            tr.CharacterFormat.TextBackgroundColor = Color.Green;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Border Around Text");
            tr.CharacterFormat.Border.BorderType = Spire.Doc.Documents.BorderStyle.Single;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Text Scale");
            tr.CharacterFormat.TextScale = 150;

            paragraph.AppendBreak(BreakType.LineBreak);
            tr = paragraph.AppendText("Character Spacing is 2 point");
            tr.CharacterFormat.CharacterSpacing = 2;

            string filePath = "CharaterFormatting.docx";
            document.SaveToFile(filePath, FileFormat.Docx);
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
