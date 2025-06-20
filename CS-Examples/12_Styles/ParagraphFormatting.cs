using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

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
            //Initialize a document
			Document document = new Document();

			//Add a section
			Section sec = document.AddSection();

			//Add a paragraph
			Paragraph para = sec.AddParagraph();

			//Append text
			para.AppendText("Paragraph Formatting");

			//Apply the Title style
			para.ApplyStyle(BuiltinStyle.Title);

			//Add a paragraph
			para = sec.AddParagraph();

			//Append text
			para.AppendText("This paragraph is surrounded with borders.");

			//Set the border type
			para.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;

			//Set the border color
			para.Format.Borders.Color = Color.Red;

			para = sec.AddParagraph();
			para.AppendText("The alignment of this paragraph is Left.");

			//Set the horizontal alignment style
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

			para = sec.AddParagraph();
			para.AppendText("The alignment of this paragraph is Center.");

			//Set the horizontal alignment style
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

			para = sec.AddParagraph();
			para.AppendText("The alignment of this paragraph is Right.");

			//Set the horizontal alignment style
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

			para = sec.AddParagraph();
			para.AppendText("The alignment of this paragraph is justified.");

			//Set the horizontal alignment style
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;

			para = sec.AddParagraph();
			para.AppendText("The alignment of this paragraph is distributed.");

			//Set the horizontal alignment style
			para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Distribute;

			para = sec.AddParagraph();
			para.AppendText("This paragraph has the gray shadow.");

			//Set the backcolor
			para.Format.BackColor = Color.Gray;

			para = sec.AddParagraph();
			para.AppendText("This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");

			//Set the indent
			para.Format.SetLeftIndent(10);
			para.Format.SetRightIndent(10);
			para.Format.SetFirstLineIndent(15);

			para = sec.AddParagraph();
			para.AppendText("The hanging indentation of this paragraph is 15pt.");
			//Negative value represents hanging indentation
			para.Format.SetFirstLineIndent(-15);

			para = sec.AddParagraph();
			para.AppendText("This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");

			//Set the spacing (in points) after the paragraph
			para.Format.AfterSpacing = 20;

			//Set the spacing (in points) before the paragraph
			para.Format.BeforeSpacing = 10;

			//Set the LineSpacingRule
			para.Format.LineSpacingRule = LineSpacingRule.AtLeast;

			//Set line spacing property of the paragraph.
			para.Format.LineSpacing = 10;

			//Save as docx file.
			string filePath = "ParagraphFormatting.docx";
			document.SaveToFile(filePath, FileFormat.Docx);

			//Dispose the document
			document.Dispose();

            //Launching the MS Word file.
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
